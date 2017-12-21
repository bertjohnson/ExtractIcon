/*
 * ExtractIcon (https://github.com/bertjohnson/extracticon)
 * 
 * Licensed according to the MIT License (http://mit-license.org/).
 * 
 * Copyright © Bert Johnson (https://bertjohnson.com/) of Allcloud Inc. (https://allcloud.com/).
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 * 
 */

using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace extracticon
{
   public partial class Program
   {
        /// <summary>
        /// Extract an icon and paint it onto a canvas, then save as PNG.
        /// </summary>
        /// <param name="args">Array of command line arguments.</param>
        static void Main(string[] args)
        {
            if (args.Length > 1)
            {
                // Parse parameters.
                string inp = args[0].Replace("file://", "").Replace("/", "\\");
                string op = args[1].Replace("file://", "").Replace("/", "\\");

                // Determine the full file path.
                StringBuilder sb = new StringBuilder(255);
                GetShortPathName(inp, sb, sb.Capacity);
                inp = sb.ToString();
                if (op.IndexOf("\\") > -1)
                {
                    string[] subfolders = op.Split('\\');
                    string folderssofar = subfolders[0];
                    for (long i = 1; i < subfolders.Length - 1; i++)
                    {
                        folderssofar += "\\" + subfolders[i];
                        if (!Directory.Exists(folderssofar))
                            Directory.CreateDirectory(folderssofar);
                    }
                }

                // Remove the output if it already exists.
                if (File.Exists(op))
                    File.Delete(op);

                try
                {
                    // Prepare blank canvas.
                    IntPtr iconHDC = CreateDC("Display", null, null, IntPtr.Zero);
                    IntPtr iconHDCDest = CreateCompatibleDC(iconHDC);
                    Bitmap iconBMP = new Bitmap(256, 256, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                    using (Graphics g = Graphics.FromImage(iconBMP))
                    {
                        g.Clear(Color.Transparent);
                    }

                    // Draw the image onto the canvas.
                    IntPtr iconHBitmap = iconBMP.GetHbitmap();
                    IntPtr iconHObj = SelectObject(iconHDCDest, iconHBitmap);
                    Guid guidImageList = new Guid("46EB5926-582E-4017-9FDF-E8998DAA0950");
                    IImageList iImageList = null;
                    IntPtr hIml = IntPtr.Zero;
                    int ret = SHGetImageList(JUMBO_SIZE, ref guidImageList, ref iImageList);
                    int ret2 = SHGetImageListHandle(JUMBO_SIZE, ref guidImageList, ref hIml);
                    DrawImage(iImageList, iconHDCDest, IconIndex(inp, true), 0, 0, ImageListDrawItemConstants.ILD_PRESERVEALPHA);
                    iconBMP.Dispose();

                    // Find the largest dimension of the copied bitmap.
                    int size = 256;
                    iconBMP = Bitmap.FromHbitmap(iconHBitmap);
                    if (CheckPixelRangeConsistency(ref iconBMP, 128, 128, 255, 255))
                    {
                        size = 128;
                        if (CheckPixelRangeConsistency(ref iconBMP, 64, 64, 127, 127))
                        {
                            size = 64;
                            if (CheckPixelRangeConsistency(ref iconBMP, 48, 48, 63, 63))
                            {
                                size = 48;
                                if (CheckPixelRangeConsistency(ref iconBMP, 32, 32, 47, 47))
                                {
                                    size = 32;
                                    if (CheckPixelRangeConsistency(ref iconBMP, 16, 16, 31, 31))
                                    {
                                        size = 16;
                                    }
                                }
                            }
                        }
                    }

                    // Resize the bitmap if needed.
                    if (size != 256)
                    {
                        iconBMP = iconBMP.Clone(new Rectangle(0, 0, size, size), iconBMP.PixelFormat);
                    }

                    // Save as a PNG.
                    iconBMP.MakeTransparent();
                    iconBMP.Save(op, System.Drawing.Imaging.ImageFormat.Png);
                    iconBMP.Dispose();

                    // Clean up handles.
                    DeleteDC(iconHDCDest);
                    DeleteObject(iconHObj);
                    DeleteObject(iconHBitmap);
                    DeleteDC(iconHDC);

                    Console.WriteLine("Success");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.ToString());
                }
            }
            else
                Console.WriteLine("Syntax: extracticon.exe [input_filename] [output_image]");
        }

        /// <summary>
        /// Check if a pixel range all has the same color.
        /// </summary>
        /// <param name="bmp">Bitmap to check.</param>
        /// <param name="startX">Starting X coordinate.</param>
        /// <param name="startY">Starting Y coordinate.</param>
        /// <param name="endX">Ending X coordinate.</param>
        /// <param name="endY">Ending Y coordinate.</param>
        /// <returns>True if all pixels in the range have the same value; otherwise, false.</returns>
        static bool CheckPixelRangeConsistency(ref Bitmap bmp, int startX, int startY, int endX, int endY)
        {
            Color finalColor = bmp.GetPixel(endX, endY);
            for (int x = 0; x <= endX; x++)
            {
                for (int y = 0; y <= endY; y++)
                {
                    if (x >= startX || y >= endY)
                    {
                        if (bmp.GetPixel(x, y) != finalColor)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Draws an image to the specified context.
        /// </summary>
        /// <param name="hdc"></param>
        /// <param name="index"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="flags"></param>
        static void DrawImage(IImageList iImageList, IntPtr hdc, int index, int x, int y, ImageListDrawItemConstants flags)
        {
            IMAGELISTDRAWPARAMS pimldp = new IMAGELISTDRAWPARAMS();
            pimldp.hdcDst = hdc;
            pimldp.cbSize = Marshal.SizeOf(pimldp.GetType());
            pimldp.i = index;
            pimldp.x = x;
            pimldp.y = y;
            pimldp.rgbFg = -1;
            pimldp.fStyle = (int)flags;
            iImageList.Draw(ref pimldp);
        }

        /// <summary>
        /// Returns the index of the icon for the specified file
        /// </summary>
        /// <param name="fileName">Filename to get icon for</param>
        /// <param name="forceLoadFromDisk">If True, then hit the disk to get the icon,
        /// otherwise only hit the disk if no cached icon is available.</param>
        /// <param name="iconState">Flags specifying the state of the icon
        /// returned.</param>
        /// <returns>Index of the icon</returns>
        static int IconIndex(string fileName, bool forceLoadFromDisk)
        {
            SHGetFileInfoConstants dwFlags = SHGetFileInfoConstants.SHGFI_SYSICONINDEX;
            int dwAttr = 0;
            SHFILEINFO shfi = new SHFILEINFO();
            uint shfiSize = (uint)Marshal.SizeOf(shfi.GetType());
            IntPtr retVal = SHGetFileInfo(fileName, dwAttr, ref shfi, shfiSize, ((uint)(dwFlags)));

            if (retVal.Equals(IntPtr.Zero))
            {
                System.Diagnostics.Debug.Assert((!retVal.Equals(IntPtr.Zero)), "Failed to get icon index");
                return 0;
            }
            else
            {
                return shfi.iIcon;
            }
        }
    }
}
    