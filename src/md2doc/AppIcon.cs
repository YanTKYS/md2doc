using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace Md2Doc;

internal static class AppIcon
{
    public static Icon Create()
    {
        using var bmp = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            // 青背景
            g.Clear(Color.FromArgb(33, 150, 243));

            // 白文字 "M"
            using var font = new Font("Segoe UI", 20, FontStyle.Bold, GraphicsUnit.Pixel);
            const string text = "M";
            var sz = g.MeasureString(text, font);
            g.DrawString(text, font, Brushes.White,
                (32 - sz.Width) / 2f, (32 - sz.Height) / 2f);
        }

        // PNG を ICO コンテナに格納（Vista 以降対応）
        using var pngStream = new MemoryStream();
        bmp.Save(pngStream, System.Drawing.Imaging.ImageFormat.Png);
        var png = pngStream.ToArray();

        using var ico = new MemoryStream();
        // ICONDIR (6 bytes)
        ico.Write(BitConverter.GetBytes((short)0));  // Reserved
        ico.Write(BitConverter.GetBytes((short)1));  // Type = ICO
        ico.Write(BitConverter.GetBytes((short)1));  // Count = 1
        // ICONDIRENTRY (16 bytes): data offset = 6 + 16 = 22
        ico.WriteByte(32);                            // Width
        ico.WriteByte(32);                            // Height
        ico.WriteByte(0);                             // ColorCount
        ico.WriteByte(0);                             // Reserved
        ico.Write(BitConverter.GetBytes((short)1));  // Planes
        ico.Write(BitConverter.GetBytes((short)32)); // BitCount
        ico.Write(BitConverter.GetBytes(png.Length));// SizeInBytes
        ico.Write(BitConverter.GetBytes(22));        // ImageOffset
        ico.Write(png);

        ico.Position = 0;
        return new Icon(ico);
    }
}
