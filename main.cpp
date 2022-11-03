#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <unknwn.h>
#include <winrt/base.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Data.Xml.Dom.h>
#include <winrt/Windows.Data.Xml.Xsl.h>
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
using winrt::Windows::Data::Xml::Dom::XmlDocument;
using winrt::Windows::Data::Xml::Dom::XmlLoadSettings;

// from ConEmu\src\ConEmu\Options.cpp
const std::unordered_map < std::wstring_view, std::array<COLORREF, 16>> CONEMU_PALETTES = { { L"Default Windows scheme", {
            0x00000000, 0x00800000, 0x00008000, 0x00808000, 0x00000080, 0x00800080, 0x00008080, 0x00c0c0c0,
            0x00808080, 0x00ff0000, 0x0000ff00, 0x00ffff00, 0x000000ff, 0x00ff00ff, 0x0000ffff, 0x00ffffff
        }
},
{ L"Babun", {
        0x001e1d1b, 0x00d28b26, 0x0014b482, 0x00d6c256, 0x007226f9, 0x00fe548c, 0x001f97fd, 0x00c6cccc,
        0x00545350, 0x00e3ad62, 0x0046ebb7, 0x00e5d894, 0x009559ff, 0x00fea0bf, 0x006cedfe, 0x00f2f8f8
    }
}, { L"Base16", {
        0x00181818, 0x00383838, 0x00d8d8d8, 0x004669a1, 0x005696dc, 0x00282828, 0x00b8b8b8, 0x00e8e8e8,
        0x00585858, 0x00c2af7c, 0x006cb5a1, 0x00b9c186, 0x004246ab, 0x00af8bba, 0x0088caf7, 0x00f8f8f8
    }
},
{ L"Cobalt2", {
        0x00000000, 0x00ff5555, 0x003bd01d, 0x00fae36a, 0x000000ff, 0x00ff55ff, 0x0009c8ed, 0x00ffffff,
        0x00555555, 0x00d26014, 0x0021de38, 0x00bbbb00, 0x00170ef4, 0x005d00ff, 0x000ae5ff, 0x00bbbbbb
    }
},
{ L"ConEmu", {
        0x00362b00, 0x00423607, 0x00808000, 0x00a48231, 0x00164bcb, 0x00b6369c, 0x00009985, 0x00d5e8ee,
        0x00a1a193, 0x00d28b26, 0x0036b64f, 0x0098a12a, 0x002f32dc, 0x008236d3, 0x000089b5, 0x00e3f6fd
    }
},
{ L"Gamma 1", {
        0x00000000, 0x00960000, 0x0000aa00, 0x00aaaa00, 0x000000aa, 0x00800080, 0x0000aaaa, 0x00c0c0c0,
        0x00808080, 0x00ff0000, 0x0000ff00, 0x00ffff00, 0x000000ff, 0x00ff00ff, 0x0000ffff, 0x00ffffff
    }
},
{ L"Monokai", {
        0x00222827, 0x009e5401, 0x0004aa74, 0x00a6831a, 0x003403a7, 0x009c5689, 0x0049b6b6, 0x00cacaca,
        0x007c7c7c, 0x00f58303, 0x0006d08d, 0x00e5c258, 0x004b04f3, 0x00b87da8, 0x0081cccc, 0x00ffffff
    }
},
{ L"Murena scheme", {
        0x00000000, 0x00644100, 0x00008000, 0x00808000, 0x00000080, 0x00800080, 0x00008080, 0x00c0c0c0,
        0x00808080, 0x00ff0000, 0x0076c587, 0x00ffff00, 0x00004bff, 0x00d78ce6, 0x0000ffff, 0x00ffffff
    }
},
{ L"PowerShell", {
        0x00000000, 0x00800000, 0x00008000, 0x00808000, 0x00000080, 0x00562401, 0x00F0EDEE, 0x00C0C0C0,
        0x00808080, 0x00ff0000, 0x0000FF00, 0x00FFFF00, 0x000000FF, 0x00FF00FF, 0x0000FFFF, 0x00FFFFFF
    }
},
{ L"Solarized", {
        0x00362b00, 0x00423607, 0x00756e58, 0x00837b65, 0x00164bcb, 0x00c4716c, 0x00969483, 0x00d5e8ee,
        0x00a1a193, 0x00d28b26, 0x00009985, 0x0098a12a, 0x002f32dc, 0x008236d3, 0x000089b5, 0x00e3f6fd
    }
},
{ L"Solarized Git", {
        0x00362b00, 0x00423607, 0x0098a12a, 0x00837b65, 0x00164bcb, 0x00756e58, 0x00969483, 0x00e3f6fd,
        0x00d5e8ee, 0x00d28b26, 0x00009985, 0x00c4716c, 0x002f32dc, 0x008236d3, 0x000089b5, 0x00a1a193
    }
},
{ L"Solarized (Luke Maciak)", {
        0x00423607, 0x00d28b26, 0x00009985, 0x000089b5, 0x002f32dc, 0x008236d3, 0x0098a12a, 0x00d5e8ee,
        0x00362b00, 0x00aaa897, 0x00756e58, 0x00837b65, 0x00004ff2, 0x00c4716c, 0x00a1a193, 0x00e3f6fd
    }
},
{ L"Solarized (John Doe)", {
        0x00362b00, 0x00423607, 0x00756e58, 0x00837b65, 0x002f32dc, 0x00c4716c, 0x00164bcb, 0x00d5e8ee,
        0x00a1a193, 0x00d28b26, 0x00009985, 0x0098a12a, 0x00969483, 0x008236d3, 0x000089b5, 0x00e3f6fd
    }
},
{ L"Solarized Light", {
        0x00d5e8ee, 0x00e3f6fd, 0x00756e58, 0x00837b65, 0x00164bcb, 0x00c4716c, 0x00969483, 0x00423607,
        0x00a1a193, 0x00d28b26, 0x00009985, 0x0098a12a, 0x002f32dc, 0x008236d3, 0x000089b5, 0x00362b00
    }
},
{ L"SolarMe", {
        0x00362b00, 0x00423607, 0x0098a12a, 0x00756e58, 0x00164bcb, 0x00542388, 0x00005875, 0x00d5e8ee,
        0x00a1a193, 0x00c4716c, 0x00009985, 0x00d28b26, 0x002f32dc, 0x008236d3, 0x000089b5, 0x00e3f6fd
    }
    },
    { L"Standard VGA", {
            0x00000000, 0x00aa0000, 0x0000aa00, 0x00aaaa00, 0x000000aa, 0x00aa00aa, 0x000055aa, 0x00aaaaaa,
            0x00555555, 0x00ff5555, 0x0055ff55, 0x00ffff55, 0x005555ff, 0x00ff55ff, 0x0055ffff, 0x00ffffff
        }
    },
    { L"tc-maxx", {
            0x00000000, RGB(11,27,59), RGB(0,128,0), RGB(0,90,135), RGB(106,7,28), RGB(128,0,128), RGB(128,128,0), RGB(40,150,177),
            RGB(128,128,128), RGB(0,0,255), RGB(0,255,0), RGB(0,215,243), RGB(190,7,23), RGB(255,0,255), RGB(255,255,0), RGB(255,255,255)
        }
    },
    { L"Terminal.app", {
            0x00000000, 0x00e12e49, 0x0024bc25, 0x00c8bb33, 0x002136c2, 0x00d338d3, 0x0027adad, 0x00cdcccb,
            0x00838381, 0x00ff3358, 0x0022e731, 0x00f0f014, 0x001f39fc, 0x00f835f9, 0x0023ecea, 0x00ebebe9
        }
    },
    { L"Tomorrow", {
            0x00ffffff, 0x00573821, 0x00004638, 0x004f4c1f, 0x00141464, 0x00542c44, 0x00005b75, 0x004c4d4d,
            0x008c908e, 0x00ae7142, 0x00008c71, 0x009f993e, 0x002928c8, 0x00a85989, 0x0000b7ea, 0x00000000
        }
    },
    { L"Tomorrow Night", {
            0x00211f1d, 0x005f5140, 0x00345e5a, 0x005b5d45, 0x00333366, 0x005d4a59, 0x003a6378, 0x00c6c8c5,
            0x00969896, 0x00bea281, 0x0068bdb5, 0x00b7ba8a, 0x006666cc, 0x00bb94b2, 0x0074c6f0, 0x00ffffff
        }
    },
    { L"Tomorrow Night Blue", {
            0x00512400, 0x00806d5d, 0x00547868, 0x0080801a, 0x00524e80, 0x00805d75, 0x00567780, 0x00ffffff,
            0x00b78572, 0x00ffdabb, 0x00a9f1d1, 0x00ffff35, 0x00a49dff, 0x00ffbbeb, 0x00adeeff, 0x00ffffff
        }
    },
    { L"Tomorrow Night Bright", {
            0x00000000, 0x006d533d, 0x0025655c, 0x00586038, 0x0029276a, 0x006c4b61, 0x00236273, 0x00eaeaea,
            0x00969896, 0x00daa67a, 0x004acab9, 0x00b1c070, 0x00534ed5, 0x00d897c3, 0x0047c5e7, 0x00ffffff
        }
    },
    { L"Tomorrow Night Eighties", {
            0x002d2d2d, 0x00664c33, 0x00386638, 0x00666633, 0x00383679, 0x00664c66, 0x004c6680, 0x00cccccc,
            0x00999999, 0x00cc9966, 0x0099cc99, 0x00cccc66, 0x007a77f2, 0x00cc99cc, 0x0099ccff, 0x00ffffff
        }
    },
    { L"Twilight", {
            0x00141414, 0x004c6acf, 0x00619083, 0x0069a8ce, 0x00a68775, 0x009d859b, 0x00b3a605, 0x00d7d7d7,
            0x00666666, 0x00a68775, 0x004da459, 0x00e6e60a, 0x00520ad8, 0x008b00e6, 0x003eeee1, 0x00e6e6e6
        }
    },
    { L"Ubuntu", { // Need to set up backround "picture" with colorfill #300A24
            0x0036342e, 0x00a46534, 0x00069a4e, 0x009a9806, 0x000000cc, 0x007b5075, 0x0000a0c4, 0x00cfd7d3,
            0x00535755, 0x00cf9f72, 0x0034e28a, 0x00e2e234, 0x002929ef, 0x00a87fad, 0x004fe9fc, 0x00eceeee
        }
    },
    { L"xterm", {
            0x00000000, 0x00ee0000, 0x0000cd00, 0x00cdcd00, 0x000000cd, 0x00cd00cd, 0x0000cdcd, 0x00e5e5e5,
            0x007f7f7f, 0x00ff5c5c, 0x0000ff00, 0x00ffff00, 0x000000ff, 0x00ff00ff, 0x0000ffff, 0x00ffffff
        }
    },
    { L"Zenburn", {
            0x003f3f3f, 0x00af6464, 0x00008000, 0x00808000, 0x00232333, 0x00aa50aa, 0x0000dcdc, 0x00ccdcdc,
            0x008080c0, 0x00ffafaf, 0x007f9f7f, 0x00d3d08c, 0x007071e3, 0x00c880c8, 0x00afdff0, 0x00ffffff
        },
    } };

constexpr auto THEME_KEY = LR"(Software\Microsoft\Windows\CurrentVersion\Themes\Personalize)";

bool is_dark_mode()
{
    std::array<char, 4> buffer = {};
    auto sz = static_cast<DWORD>(buffer.size() * sizeof(char));
    RegGetValueW(
        HKEY_CURRENT_USER,
        THEME_KEY,
        L"AppsUseLightTheme",
        RRF_RT_REG_DWORD,
        nullptr,
        buffer.data(),
        &sz);

    return !*reinterpret_cast<const uint32_t*>(buffer.data());
}


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    constexpr bool use_dedicated_tmp_dir = false;
    std::wstring conemu_path = LR"d(C:\Program Files\ConEmu\ConEmu64.exe)d";
    std::wstring light_theme_name = L"Tomorrow";
    std::wstring dark_theme_name = L"Tomorrow Night";

    int argc = 0;
    const auto argv = CommandLineToArgvW(GetCommandLineW(), &argc);

    if (argc == 4)
    {
        conemu_path = argv[1];
        light_theme_name = argv[2];
        dark_theme_name = argv[3];
    }

    const auto& light_theme_colors = CONEMU_PALETTES.at(light_theme_name);
    const auto& dark_theme_colors = CONEMU_PALETTES.at(dark_theme_name);

    char tmpdir[MAX_PATH];
    try
    {
        if (use_dedicated_tmp_dir)
        {
            tmpnam_s(tmpdir);
            fs::create_directory(tmpdir);
        }
        else
        {
            const auto tmpdir_str = fs::temp_directory_path().string();
            tmpdir_str.copy(tmpdir, tmpdir_str.size());
            tmpdir[tmpdir_str.size()] = '\0';
        }
    }
    catch (const std::exception&)
    {
        return 1;
    }

    std::wstring conemu_config_path;
    conemu_config_path.resize(MAX_PATH);
    SHGetSpecialFolderPathW({}, conemu_config_path.data(), CSIDL_APPDATA, false);
    conemu_config_path.resize(wcslen(conemu_config_path.c_str()));
    conemu_config_path += L"\\ConEmu.xml";

    const auto light_bat = (fs::path{ tmpdir } / "light.bat").wstring();
    const auto dark_bat = (fs::path{ tmpdir } / "dark.bat").wstring();
    const auto create_batch = [&](const auto light, const auto path)
    {
        const auto contents = std::format(LR"d(ConEmuC -GuiMacro SetOption("Scheme", "<{}>")
ConEmuC -GuiMacro WindowMinimize
ConEmuC -GuiMacro Close(0,1))d", light ? light_theme_name : dark_theme_name);
        std::wofstream f{ path };
        f.write(contents.data(), contents.size());
    };

    create_batch(true, light_bat);
    create_batch(false, dark_bat);

    HKEY key{};

    RegOpenKeyExW(HKEY_CURRENT_USER, THEME_KEY, 0, KEY_NOTIFY, &key);
    HANDLE event = CreateEventW(nullptr, true, false, nullptr);
    RegNotifyChangeKeyValue(key, true, REG_NOTIFY_CHANGE_LAST_SET, event, true);

    bool is_dark = is_dark_mode();
    std::wstring run_bat;

    const auto update_settings = [&](const auto& colors)
    {
        XmlDocument settings;
        XmlLoadSettings load_settings;
        load_settings.ValidateOnParse(false);
        load_settings.ProhibitDtd(false);
        std::wfstream f{ conemu_config_path };
        f.seekg(0, std::ios::end);
        const size_t file_size = f.tellg();
        f.seekg(0, std::ios::beg);
        std::wstring settings_xml;
        settings_xml.resize(file_size);
        f.read(settings_xml.data(), file_size);
        // https://learn.microsoft.com/en-us/troubleshoot/developer/dotnet/framework/general/xml-parser-invalid-character
        for(size_t i = 0; i < file_size; ++i)
            assert(settings_xml[i] <= 0x7F);
        settings.LoadXml(settings_xml, load_settings);
        for (auto& setting : settings.SelectNodes(L"key/key/key/value"))
        {
            auto& setting_attrs = setting.Attributes();
            const auto setting_name = setting_attrs.GetNamedItem(L"name").InnerText();
            const auto snv = std::wstring_view{ setting_name.c_str() };
            constexpr auto name_prefix = std::wstring_view{ L"ColorTable" };
            if (!snv.starts_with(name_prefix))
                continue;

            int color_id = 0;
            swscanf_s(snv.data() + name_prefix.size(), L"%d", &color_id);
            wchar_t buf[16];
            swprintf_s(buf, L"%08x", colors[color_id]);
            setting_attrs.GetNamedItem(L"data").InnerText(buf);
        }

        const auto updated_settings = settings.GetXml();
        f.write(updated_settings.c_str(), updated_settings.size());
    };

    update_settings(light_theme_colors);

    while (WaitForSingleObject(event, INFINITE) != WAIT_FAILED)
    {
        const bool new_dark = is_dark_mode();
        ResetEvent(event);
        RegNotifyChangeKeyValue(key, true, REG_NOTIFY_CHANGE_LAST_SET, event, true);

        if (new_dark == is_dark)
            continue;

        run_bat.clear();
        std::format_to(std::back_inserter(run_bat), L"\"{}\" {}", conemu_path, new_dark ? dark_bat : light_bat);
        _wsystem(run_bat.c_str());
        is_dark = new_dark;
    }

    return 0;
}
