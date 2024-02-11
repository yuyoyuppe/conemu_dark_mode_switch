#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <unknwn.h>
#include <shellapi.h>
#include <Shlobj_core.h>

#include <array>
#include <cstdint>
#include <cstdio>
#include <string>
#include <format>
#include <filesystem>
#include <fstream>
#include <string_view>
#include <unordered_map>

namespace fs = std::filesystem;

namespace {
constexpr auto THEME_KEY = LR"(Software\Microsoft\Windows\CurrentVersion\Themes\Personalize)";
bool           is_dark_mode() {
    std::array<char, 4> buffer = {};
    auto                sz     = static_cast<DWORD>(buffer.size() * sizeof(char));
    RegGetValueW(HKEY_CURRENT_USER, THEME_KEY, L"AppsUseLightTheme", RRF_RT_REG_DWORD, nullptr, buffer.data(), &sz);

    return !*reinterpret_cast<const uint32_t *>(buffer.data());
}
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    constexpr bool use_dedicated_tmp_dir = false;
    std::wstring   conemu_path           = LR"d(C:\Program Files\ConEmu\ConEmu64.exe)d";
    std::wstring   light_theme_name      = L"Tomorrow";
    std::wstring   dark_theme_name       = L"Tomorrow Night";

    int        argc = 0;
    const auto argv = CommandLineToArgvW(GetCommandLineW(), &argc);

    if(argc == 4) {
        conemu_path      = argv[1];
        light_theme_name = argv[2];
        dark_theme_name  = argv[3];
    }

    char tmpdir[MAX_PATH];
    try {
        if(use_dedicated_tmp_dir) {
            tmpnam_s(tmpdir);
            fs::create_directory(tmpdir);
        } else {
            const auto tmpdir_str = fs::temp_directory_path().string();
            tmpdir_str.copy(tmpdir, tmpdir_str.size());
            tmpdir[tmpdir_str.size()] = '\0';
        }
    } catch(const std::exception &) { return 1; }

    const auto light_bat    = (fs::path{tmpdir} / "light.bat").wstring();
    const auto dark_bat     = (fs::path{tmpdir} / "dark.bat").wstring();
    const auto create_batch = [&](const auto light, const auto path) {
        const auto contents = std::format(
          LR"d(ConEmuC -GuiMacro palette 1 "<{}>")
ConEmuC -GuiMacro WindowMinimize
ConEmuC -GuiMacro Close(0,1))d",
          light ? light_theme_name : dark_theme_name);
        std::wofstream f{path};
        f.write(contents.data(), contents.size());
    };

    create_batch(true, light_bat);
    create_batch(false, dark_bat);

    HKEY key{};

    RegOpenKeyExW(HKEY_CURRENT_USER, THEME_KEY, 0, KEY_NOTIFY, &key);
    HANDLE event = CreateEventW(nullptr, true, false, nullptr);
    RegNotifyChangeKeyValue(key, true, REG_NOTIFY_CHANGE_LAST_SET, event, true);

    bool         is_dark = is_dark_mode();
    std::wstring run_bat;
    while(WaitForSingleObject(event, INFINITE) != WAIT_FAILED) {
        const bool new_dark = is_dark_mode();
        ResetEvent(event);
        RegNotifyChangeKeyValue(key, true, REG_NOTIFY_CHANGE_LAST_SET, event, true);

        if(new_dark == is_dark)
            continue;

        run_bat.clear();
        std::format_to(std::back_inserter(run_bat), L"\"{}\" {}", conemu_path, new_dark ? dark_bat : light_bat);
        _wsystem(run_bat.c_str());
        is_dark = new_dark;
    }

    return 0;
}
