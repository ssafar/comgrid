using System;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Forms;

// `cgr snap RANGE file.png`: the range as Excel renders it on screen, via the clipboard
// (Range.CopyPicture), so an agent can see formatting, charts and layout.
static class Snap
{
    const int xlScreen = 1, xlBitmap = 2;

    public static void Picture(dynamic ws, dynamic rng, string path, Out o)
    {
        string full = Path.GetFullPath(path);
        for (int tries = 1; ; tries++)   // fails now and then while the clipboard is busy
        {
            try { rng.CopyPicture(xlScreen, xlBitmap); break; }
            catch (System.Runtime.InteropServices.COMException) when (tries < 5) { Thread.Sleep(300); }
        }
        System.Drawing.Image img = null;
        for (int tries = 0; tries < 20 && img == null; tries++)   // Excel fills the clipboard asynchronously
        {
            if (Clipboard.ContainsImage()) img = Clipboard.GetImage();
            else Thread.Sleep(100);
        }
        if (img == null) throw new CgrError("Excel put no picture on the clipboard");
        string size = $"{img.Width}x{img.Height}";
        using (img) img.Save(full, ImageFormat.Png);
        o.Line($"saved {Xl.FullOf(rng)} as {full} ({size}); note: this replaced the clipboard");
    }
}
