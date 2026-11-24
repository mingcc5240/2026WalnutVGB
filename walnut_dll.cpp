#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <mmsystem.h>
#include <stdint.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>

#pragma comment(lib, "winmm.lib") // Windows 오디오 출력용 라이브러리 자동 링크

// 1. MiniGB APU (C 언어 헤더이므로 extern "C"로 감쌈)
extern "C" {
#include "minigb_apu.h"
}

// 2. Walnut-CGB 설정 매크로
#define ENABLE_LCD 1
#define ENABLE_SOUND 1
#define WALNUT_GB_32BIT_DMA 1

// --- 오디오 레지스터 읽기/쓰기 콜백 (minigb_apu 연동) ---
// --- 전역 에뮬레이터 및 오디오 상태 ---

static struct minigb_apu_ctx apu;
uint8_t audio_read(const uint16_t addr) {
    return minigb_apu_audio_read(&apu, addr);
}

void audio_write(const uint16_t addr, const uint8_t val) {
    minigb_apu_audio_write(&apu, addr, val);
}


//#pragma warning(push)
//#pragma warning(disable: 4244 4018 4146 4996)
#include "walnut_cgb.h"
//#pragma warning(pop)

static struct gb_s gb;

#define EXPORT extern "C" __declspec(dllexport)


static uint8_t* rom_data = nullptr;
static size_t rom_size = 0;
static uint8_t* cart_ram = nullptr;
static size_t cart_ram_size = 0;

static uint32_t frame_buffer[LCD_HEIGHT * LCD_WIDTH];

// 흑백(DMG) 기본 4색 팔레트
static const uint32_t dmg_palette[] = {
    0xFF9BBC0F, 0xFF8BAC0F, 0xFF306230, 0xFF0F380F
};

// --- Windows waveOut 더블 버퍼링 오디오 설정 ---
#define AUDIO_BUF_COUNT 4

static HWAVEOUT hWaveOut = nullptr;
static WAVEHDR waveHeaders[AUDIO_BUF_COUNT];


//#define SAMPLES_PER_FRAME (AUDIO_SAMPLE_RATE / 60) // 약 735~800 샘플 (스테레오 기준 x2)
//static int16_t audioBuffers[AUDIO_BUF_COUNT][SAMPLES_PER_FRAME * 2]; // 16-bit Stereo
// SAMPLES_PER_FRAME 매크로 삭제 후 아래처럼 변경
static int16_t audioBuffers[AUDIO_BUF_COUNT][AUDIO_SAMPLES_TOTAL];
//

static int currentAudioBuf = 0;



// --- Walnut-CGB 필수 콜백 함수들 ---

uint8_t gb_rom_read(struct gb_s* gb_ctx, const uint_fast32_t addr) {
    return (rom_data && addr < rom_size) ? rom_data[addr] : 0xFF;
}

uint16_t gb_rom_read_16bit(struct gb_s* gb_ctx, const uint_fast32_t addr) {
    if (rom_data && addr + 1 < rom_size) {
        return *(uint16_t*)(rom_data + addr);
    }
    return 0xFFFF;
}

uint32_t gb_rom_read_32bit(struct gb_s* gb_ctx, const uint_fast32_t addr) {
    if (rom_data && addr + 3 < rom_size) {
        return *(uint32_t*)(rom_data + addr);
    }
    return 0xFFFFFFFF;
}

uint8_t gb_cart_ram_read(struct gb_s* gb_ctx, const uint_fast32_t addr) {
    return (cart_ram && addr < cart_ram_size) ? cart_ram[addr] : 0xFF;
}

void gb_cart_ram_write(struct gb_s* gb_ctx, const uint_fast32_t addr, const uint8_t val) {
    if (cart_ram && addr < cart_ram_size) {
        cart_ram[addr] = val;
    }
}

void gb_error(struct gb_s* gb_ctx, const enum gb_error_e err, const uint16_t addr) {}


// --- LCD 렌더링 콜백 (RGB565 -> 32비트 ARGB 변환) ---

static inline uint32_t rgb565_to_argb(uint16_t color) {
    // Walnut-CGB의 fixPalette는 기본 RGB565 포맷
    uint8_t r = (color >> 11) & 0x1F;
    uint8_t g = (color >> 5) & 0x3F;
    uint8_t b = color & 0x1F;

    r = (r << 3) | (r >> 2);
    g = (g << 2) | (g >> 4);
    b = (b << 3) | (b >> 2);

    return 0xFF000000 | (r << 16) | (g << 8) | b;
}

void lcd_draw_line(struct gb_s* gb_ctx, const uint8_t* pixels, const uint_fast8_t line) {
    if (line >= LCD_HEIGHT) return;

    uint32_t* dst = &frame_buffer[line * LCD_WIDTH];

    if (gb_ctx->cgb.cgbMode) {
        for (int x = 0; x < LCD_WIDTH; x++) {
            dst[x] = rgb565_to_argb(gb_ctx->cgb.fixPalette[pixels[x]]);
        }
    } else {
        for (int x = 0; x < LCD_WIDTH; x++) {
            dst[x] = dmg_palette[pixels[x] & 0x03];
        }
    }
}

// --- 오디오 장치 초기화 및 해제 ---

void InitAudioDevice() {
    if (hWaveOut) return;

    WAVEFORMATEX wfx = {};
    wfx.wFormatTag = WAVE_FORMAT_PCM;
    wfx.nChannels = 2; // 스테레오
    wfx.nSamplesPerSec = AUDIO_SAMPLE_RATE;
    wfx.wBitsPerSample = 16;
    wfx.nBlockAlign = (wfx.nChannels * wfx.wBitsPerSample) / 8;
    wfx.nAvgBytesPerSec = wfx.nSamplesPerSec * wfx.nBlockAlign;

    if (waveOutOpen(&hWaveOut, WAVE_MAPPER, &wfx, 0, 0, CALLBACK_NULL) == MMSYSERR_NOERROR) {
        memset(waveHeaders, 0, sizeof(waveHeaders));
        for (int i = 0; i < AUDIO_BUF_COUNT; i++) {
            waveHeaders[i].lpData = (LPSTR)audioBuffers[i];
            waveHeaders[i].dwBufferLength = sizeof(audioBuffers[i]);
            waveOutPrepareHeader(hWaveOut, &waveHeaders[i], sizeof(WAVEHDR));
        }
    }
}

void CloseAudioDevice() {
    if (hWaveOut) {
        waveOutReset(hWaveOut);
        for (int i = 0; i < AUDIO_BUF_COUNT; i++) {
            waveOutUnprepareHeader(hWaveOut, &waveHeaders[i], sizeof(WAVEHDR));
        }
        waveOutClose(hWaveOut);
        hWaveOut = nullptr;
    }
}

// --- Visual Basic 호출용 외부 API ---

EXPORT int __stdcall GB_Init(const char* rom_path) {
    if (rom_data) { free(rom_data); rom_data = nullptr; }
    if (cart_ram) { free(cart_ram); cart_ram = nullptr; }

    FILE* fp = nullptr;
    if (fopen_s(&fp, rom_path, "rb") != 0 || !fp) return -1;

    fseek(fp, 0, SEEK_END);
    rom_size = ftell(fp);
    fseek(fp, 0, SEEK_SET);

    rom_data = (uint8_t*)malloc(rom_size);
    if (!rom_data) { fclose(fp); return -2; }

    fread(rom_data, 1, rom_size, fp);
    fclose(fp);

    // Walnut-CGB 초기화 (16비트/32비트 ROM 읽기 콜백 포함)
    enum gb_init_error_e ret = gb_init(
        &gb,
        &gb_rom_read,
        &gb_rom_read_16bit,
        &gb_rom_read_32bit,
        &gb_cart_ram_read,
        &gb_cart_ram_write,
        &gb_error,
        nullptr
    );
    if (ret != GB_INIT_NO_ERROR) return -3;

    cart_ram_size = gb_get_save_size(&gb);
    if (cart_ram_size > 0) {
        cart_ram = (uint8_t*)calloc(1, cart_ram_size);
    }

    // LCD 및 APU 초기화
    gb_init_lcd(&gb, &lcd_draw_line);
    minigb_apu_audio_init(&apu);
    InitAudioDevice();

    memset(frame_buffer, 0, sizeof(frame_buffer));
    return 0;
}

EXPORT void __stdcall GB_RunFrame() {
    if (!rom_data) return;

    // 고속 Dual-Fetch 모드로 1프레임 실행
    gb_run_frame_dualfetch(&gb);

    // 1프레임 분량의 오디오 샘플 생성 및 재생 큐 전송
    if (hWaveOut) {
        WAVEHDR* hdr = &waveHeaders[currentAudioBuf];
        // 이전 재생이 완료된 버퍼일 경우에만 새 오디오 데이터 채우기
        if (!(hdr->dwFlags & WHDR_INQUEUE)) {
            minigb_apu_audio_callback(&apu, (int16_t*)hdr->lpData);
            waveOutWrite(hWaveOut, hdr, sizeof(WAVEHDR));
            currentAudioBuf = (currentAudioBuf + 1) % AUDIO_BUF_COUNT;
        }
    }
}

EXPORT void __stdcall GB_SetJoypad(int up, int down, int left, int right, int a, int b, int select, int start) {
    gb.direct.joypad_bits.up = up ? 0 : 1;
    gb.direct.joypad_bits.down = down ? 0 : 1;
    gb.direct.joypad_bits.left = left ? 0 : 1;
    gb.direct.joypad_bits.right = right ? 0 : 1;
    gb.direct.joypad_bits.a = a ? 0 : 1;
    gb.direct.joypad_bits.b = b ? 0 : 1;
    gb.direct.joypad_bits.select = select ? 0 : 1;
    gb.direct.joypad_bits.start = start ? 0 : 1;
}

EXPORT void* __stdcall GB_GetFrameBufferPtr() {
    return (void*)frame_buffer;
}

EXPORT void __stdcall GB_Reset() {
    if (rom_data) {
        gb_reset(&gb);
        minigb_apu_audio_init(&apu);
    }
}

EXPORT void __stdcall GB_Cleanup() {
    CloseAudioDevice();
    if (rom_data) { free(rom_data); rom_data = nullptr; }
    if (cart_ram) { free(cart_ram); cart_ram = nullptr; }
    rom_size = 0;
    cart_ram_size = 0;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
    if (ul_reason_for_call == DLL_PROCESS_DETACH) GB_Cleanup();
    return TRUE;
}