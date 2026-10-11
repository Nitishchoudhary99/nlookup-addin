using System;
using System.Collections.Generic;
using System.Globalization;

namespace SmartPro
{
    // Colour arithmetic. Colours are Excel colour numbers: red + 256 * green + 65536 * blue.
    internal static class ColorMath
    {
        public static int Rgb(int r, int g, int b)
        {
            if (r < 0) r = 0;
            if (r > 255) r = 255;
            if (g < 0) g = 0;
            if (g > 255) g = 255;
            if (b < 0) b = 0;
            if (b > 255) b = 255;
            return r | (g << 8) | (b << 16);
        }

        public static int R(int c) { return c & 255; }
        public static int G(int c) { return (c >> 8) & 255; }
        public static int B(int c) { return (c >> 16) & 255; }

        // a mixed with b: t = 0 gives a, t = 1 gives b
        public static int Mix(int a, int b, double t)
        {
            if (t < 0) t = 0;
            if (t > 1) t = 1;
            int r = (int)Math.Round(R(a) + (R(b) - R(a)) * t);
            int g = (int)Math.Round(G(a) + (G(b) - G(a)) * t);
            int bl = (int)Math.Round(B(a) + (B(b) - B(a)) * t);
            return Rgb(r, g, bl);
        }

        // how bright the colour looks, 0 (black) to 1 (white)
        public static double Lum(int c)
        {
            return (0.299 * R(c) + 0.587 * G(c) + 0.114 * B(c)) / 255.0;
        }

        // a text colour that can be read on this background
        public static int Ink(int background)
        {
            return Lum(background) > 0.62 ? Rgb(15, 23, 42) : Rgb(255, 255, 255);
        }

        public static int Darker(int c, double t) { return Mix(c, 0, t); }
        public static int Lighter(int c, double t) { return Mix(c, 0xFFFFFF, t); }

        // #rrggbb
        public static string Hex(int c)
        {
            return "#" + R(c).ToString("x2", CultureInfo.InvariantCulture) + G(c).ToString("x2", CultureInfo.InvariantCulture) + B(c).ToString("x2", CultureInfo.InvariantCulture);
        }

        static int HexDigit(char ch)
        {
            if (ch >= '0' && ch <= '9') return ch - '0';
            if (ch >= 'a' && ch <= 'f') return ch - 'a' + 10;
            if (ch >= 'A' && ch <= 'F') return ch - 'A' + 10;
            return -1;
        }

        // "#2563EB" or "2563eb" to a colour number; -1 when the text is not a colour
        public static int FromHex(string s)
        {
            if (s == null) return -1;
            s = s.Trim();
            if (s.Length > 0 && s[0] == '#') s = s.Substring(1);
            if (s.Length != 6) return -1;
            int[] d = new int[6];
            for (int i = 0; i < 6; i++)
            {
                d[i] = HexDigit(s[i]);
                if (d[i] < 0) return -1;
            }
            return Rgb(d[0] * 16 + d[1], d[2] * 16 + d[3], d[4] * 16 + d[5]);
        }

        // the colour of a card or chart: its own colour, one of the six theme colours, or the automatic one for its position
        public static int Pick(ThemeSpec t, int code, int position)
        {
            if (code >= CardSpec.CustomColor) return (code - CardSpec.CustomColor) & 0xFFFFFF;
            if (code >= 0 && code < 6) return t.Accent[code];
            int p = position % 6;
            if (p < 0) p += 6;
            return t.Accent[p];
        }
    }

    // =====================================================================
    //  DASHBOARD THEMES
    //  One row for each theme: name, L (light) or D (dark), then the colours as #RRGGBB:
    //  page, header band, header text, card, text, soft text, grid lines, and six chart colours.
    //  Add a row to add a theme. The first twelve keep the order of the older "Style" setting.
    // =====================================================================
    internal static class DashThemes
    {
        static readonly string[][] Rows =
        {
            new[] { "Navy Classic", "L", "#F1F5F9", "#0F172A", "#FFFFFF", "#FFFFFF", "#0F172A", "#64748B", "#E2E8F0", "#2563EB", "#0D9488", "#D97706", "#7C3AED", "#E11D48", "#0891B2" },
            new[] { "Corporate Blue", "L", "#EFF6FF", "#1E3A8A", "#FFFFFF", "#FFFFFF", "#172554", "#64748B", "#DBEAFE", "#2563EB", "#0E7490", "#4F46E5", "#0284C7", "#475569", "#7C3AED" },
            new[] { "Emerald", "L", "#F0FDF4", "#064E3B", "#FFFFFF", "#FFFFFF", "#052E16", "#5B7065", "#D1FAE5", "#059669", "#0D9488", "#CA8A04", "#166534", "#475569", "#0891B2" },
            new[] { "Royal Purple", "L", "#FAF5FF", "#3B0764", "#FFFFFF", "#FFFFFF", "#2E1065", "#6B6280", "#EDE9FE", "#7E22CE", "#BE185D", "#4F46E5", "#A21CAF", "#475569", "#0D9488" },
            new[] { "Sunset", "L", "#FFF7ED", "#7C2D12", "#FFFFFF", "#FFFFFF", "#431407", "#78716C", "#FED7AA", "#EA580C", "#DC2626", "#CA8A04", "#BE123C", "#57534E", "#9333EA" },
            new[] { "Charcoal", "L", "#F3F4F6", "#1F2937", "#FFFFFF", "#FFFFFF", "#111827", "#6B7280", "#E5E7EB", "#374151", "#0D9488", "#6B7280", "#4B5563", "#111827", "#D97706" },
            new[] { "Steel", "L", "#F1F5F9", "#334155", "#FFFFFF", "#FFFFFF", "#1E293B", "#64748B", "#E2E8F0", "#475569", "#0369A1", "#64748B", "#0E7490", "#1E293B", "#B45309" },
            new[] { "Teal Ocean", "L", "#F0FDFA", "#134E4A", "#FFFFFF", "#FFFFFF", "#042F2E", "#5E7573", "#CCFBF1", "#0D9488", "#0891B2", "#0284C7", "#059669", "#475569", "#7C3AED" },
            new[] { "Crimson", "L", "#FEF2F2", "#7F1D1D", "#FFFFFF", "#FFFFFF", "#450A0A", "#78716C", "#FECACA", "#DC2626", "#E11D48", "#EA580C", "#B45309", "#57534E", "#7C3AED" },
            new[] { "Indigo Night", "L", "#EEF2FF", "#312E81", "#FFFFFF", "#FFFFFF", "#1E1B4B", "#64668A", "#E0E7FF", "#4F46E5", "#7C3AED", "#0284C7", "#C026D3", "#475569", "#0D9488" },
            new[] { "Forest Gold", "L", "#F7FEE7", "#14532D", "#FFFFFF", "#FFFFFF", "#1A2E05", "#667358", "#D9F99D", "#16A34A", "#65A30D", "#D97706", "#A16207", "#57534E", "#0D9488" },
            new[] { "Rose", "L", "#FFF1F2", "#831843", "#FFFFFF", "#FFFFFF", "#4C0519", "#7A6470", "#FECDD3", "#DB2777", "#E11D48", "#9333EA", "#0284C7", "#475569", "#D97706" },
            new[] { "Clean White", "L", "#FFFFFF", "#FFFFFF", "#0F172A", "#F8FAFC", "#0F172A", "#64748B", "#E2E8F0", "#2563EB", "#0D9488", "#F59E0B", "#7C3AED", "#E11D48", "#0891B2" },
            new[] { "Power Blue", "L", "#F3F2F1", "#243A5E", "#FFFFFF", "#FFFFFF", "#252423", "#605E5C", "#E1DFDD", "#118DFF", "#12239E", "#E66C37", "#6B007B", "#E044A7", "#744EC2" },
            new[] { "Studio Bright", "L", "#F8F9FA", "#1A73E8", "#FFFFFF", "#FFFFFF", "#202124", "#5F6368", "#DADCE0", "#4285F4", "#EA4335", "#FBBC04", "#34A853", "#FF6D01", "#46BDC6" },
            new[] { "Mono Ink", "L", "#FAFAFA", "#111827", "#FFFFFF", "#FFFFFF", "#111827", "#6B7280", "#E5E7EB", "#111827", "#4B5563", "#9CA3AF", "#374151", "#6B7280", "#1F2937" },
            new[] { "Sand", "L", "#FAF7F2", "#44403C", "#FFFFFF", "#FFFFFF", "#292524", "#78716C", "#E7E5E4", "#B45309", "#0F766E", "#A16207", "#9A3412", "#57534E", "#1D4ED8" },
            new[] { "Sky", "L", "#F0F9FF", "#075985", "#FFFFFF", "#FFFFFF", "#082F49", "#5B7A8C", "#BAE6FD", "#0284C7", "#0EA5E9", "#6366F1", "#14B8A6", "#F59E0B", "#EC4899" },
            new[] { "Midnight", "D", "#0B1220", "#020617", "#F8FAFC", "#131C31", "#E5E7EB", "#94A3B8", "#26334D", "#60A5FA", "#34D399", "#FBBF24", "#A78BFA", "#FB7185", "#22D3EE" },
            new[] { "Slate Dark", "D", "#0F172A", "#020617", "#F1F5F9", "#1E293B", "#F1F5F9", "#94A3B8", "#334155", "#38BDF8", "#818CF8", "#34D399", "#FBBF24", "#FB7185", "#C084FC" },
            new[] { "Carbon", "D", "#111111", "#000000", "#F5F5F5", "#1C1C1C", "#F5F5F5", "#A3A3A3", "#2E2E2E", "#F97316", "#22C55E", "#EAB308", "#38BDF8", "#F43F5E", "#A855F7" },
            new[] { "Deep Ocean", "D", "#06202B", "#03141B", "#E0F2FE", "#0B2E3C", "#E0F2FE", "#8AB4C4", "#164152", "#22D3EE", "#2DD4BF", "#FBBF24", "#60A5FA", "#F472B6", "#A3E635" },
            new[] { "Aubergine", "D", "#1A1025", "#0F0718", "#F3E8FF", "#251736", "#F3E8FF", "#B7A3CC", "#3A2752", "#C084FC", "#F472B6", "#FBBF24", "#60A5FA", "#34D399", "#FB923C" },
            new[] { "Forest Night", "D", "#0C1A12", "#06110B", "#ECFDF5", "#12261A", "#ECFDF5", "#94B8A3", "#1F3D2B", "#34D399", "#A3E635", "#FBBF24", "#38BDF8", "#F472B6", "#FB923C" },
            new[] { "Graphite Gold", "D", "#18181B", "#09090B", "#FAFAFA", "#27272A", "#FAFAFA", "#A1A1AA", "#3F3F46", "#FACC15", "#F59E0B", "#A3E635", "#38BDF8", "#F472B6", "#C084FC" },
            new[] { "Neon", "D", "#0A0A0F", "#000000", "#FFFFFF", "#14141F", "#FFFFFF", "#9CA3AF", "#262637", "#00E5FF", "#FF2E93", "#B6FF00", "#FFB800", "#8B5CF6", "#FF6B00" }
        };

        public static int Count { get { return Rows.Length; } }

        public static string NameOf(int i)
        {
            if (i < 0 || i >= Rows.Length) i = 0;
            return Rows[i][0];
        }

        public static bool IsDark(int i)
        {
            if (i < 0 || i >= Rows.Length) i = 0;
            return Rows[i][1] == "D";
        }

        public static string[] Names()
        {
            string[] n = new string[Rows.Length];
            for (int i = 0; i < n.Length; i++) n[i] = Rows[i][0] + (Rows[i][1] == "D" ? "  (dark)" : "");
            return n;
        }

        static int C(string hex)
        {
            int c = ColorMath.FromHex(hex);
            return c < 0 ? 0 : c;
        }

        public static ThemeSpec Get(int i)
        {
            if (i < 0 || i >= Rows.Length) i = 0;
            string[] r = Rows[i];
            ThemeSpec t = new ThemeSpec();
            t.Name = r[0];
            t.Dark = r[1] == "D";
            t.Page = C(r[2]);
            t.Band = C(r[3]);
            t.BandText = C(r[4]);
            t.Card = C(r[5]);
            t.Text = C(r[6]);
            t.Muted = C(r[7]);
            t.Grid = C(r[8]);
            t.Accent = new int[6];
            for (int k = 0; k < 6; k++) t.Accent[k] = C(r[9 + k]);
            t.Font = "Segoe UI";
            return t;
        }

        // -1 when no theme has this name
        public static int IndexOf(string name)
        {
            for (int i = 0; i < Rows.Length; i++)
            {
                if (string.Equals(Rows[i][0], name, StringComparison.OrdinalIgnoreCase)) return i;
            }
            return -1;
        }

        public static readonly string[] Fonts = { "Segoe UI", "Calibri", "Arial", "Verdana", "Tahoma", "Trebuchet MS", "Century Gothic", "Georgia", "Cambria" };

        public static readonly string[] CardStyles =
        {
            "Filled colour", "Light with a top line", "Light with a side bar", "Outline", "Soft tint", "Gradient"
        };
    }
}
