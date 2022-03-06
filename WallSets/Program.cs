using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using PexelsDotNetSDK.Api;
using System.Net.Http;
using System.Text.RegularExpressions;   // for Regex

namespace WallSets
{
    class Program
    {
        private const int SPI_SETDESKWALLPAPER = 20;
        private const int SPIF_UPDATEINIFILE = 0x01;
        private const int SPIF_SENDWININICHANGE = 0x02;
        static void Main(string[] args)
        {
            var (name, url) = GetBingWallPaper();
            var savedTo = Program.DownloadWallpaper(name, url);
            Program.Set(savedTo);
        }

        /// <summary>
        /// Get top wallpaper from /r/wallpaper subreddit
        /// </summary>
        /// <returns>The title and url of image</returns>
        public static (string, string) GetRedditWallPaper()
        {
            // gets the current top post from r/wallpapers
            var client = new HttpClient();
            var response = client.GetAsync("https://www.reddit.com/r/wallpaper/top.json?sort=top&t=day&limit=3").Result;
            var json = response.Content.ReadAsStringAsync().Result;
            var data = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(json);
            var children = data.data.children;
            string pattern = @"[\[\(\{](\d+)[\sxX\*]+(\d+)[\)\}\]]";
            int dimensions = 0;
            int i = 0;
            int topscore = data.data.children[0].data.score;
            int max = 0;
            foreach (var child in children)
            {
                var title = child.data.title.Value;
                var resolution = Regex.Match(title, pattern);
                var popularity = child.data.score.Value;
                if (popularity / topscore > 0.75)
                {
                    var pixels = Int32.Parse(resolution.Groups[1].Value) * Int32.Parse(resolution.Groups[2].Value);
                    if (pixels > dimensions)
                    {
                        dimensions = pixels;
                        max = i;
                    }
                    i++;
                }
            }
            var url = children[max].data.url.Value;
            var name = children[max].data.title.Value;
            return (name, url);
        }

        /// <summary>
        /// Gets daily bing wallpaper
        /// </summary>
        /// <returns>The title and url of image</returns>        
        public static (string, string) GetBingWallPaper()
        {
            var baseURL = "https://www.bing.com/";
            var client = new HttpClient();
            var response = client.GetAsync("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US").Result;
            var json = response.Content.ReadAsStringAsync().Result;
            var data = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(json);
            var url = data.images[0].url.Value.Replace("1920x1080", "UHD");
            var urlBase = data.images[0].urlbase.Value;
            var name = urlBase.Substring(urlBase.LastIndexOf('.') + 1, urlBase.LastIndexOf('_') + 1);
            return (name, baseURL + url);
        }

        /// <summary>
        /// Downloads the wallpaper from the given url and saves it to the given path
        /// </summary>
        /// <param name="title">The name of the wallpaper</param>
        /// <param name="photoUrl">The url of the wallpaper</param>
        /// <returns>The path to the saved wallpaper</returns>
        public static string DownloadWallpaper(string title, string photoUrl)
        {
            // Save image from url
            var fileName = @"Wallpapers\" + title + ".jpg";
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), fileName);
            using (var client = new System.Net.WebClient())
            {
                client.DownloadFile(photoUrl, filePath);
            }
            return filePath;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

        /// <summary>
        /// Fetches the desktop wallpaper from Pexels.com
        /// </summary>
        public static void FetchWallpaper()
        {
            var pexelClient = new PexelsClient("pexelkeys");
            var result = pexelClient.SearchPhotosAsync("nature", "landscape", "large", "", "", 1, 1);
            result.Wait();
            var photo = result.Result;
            var photoId = photo.photos[0].id;
            pexelClient.GetPhotoAsync(photoId).Wait();
            var photoResult = pexelClient.GetPhotoAsync(photoId).Result;
            var photoUrl = photoResult.source.landscape;
        }

        public enum Style : int
        {
            /// <summary>
            /// Current windows wallpaper style
            /// </summary>
            Current = -1,
            /// <summary>
            /// Centered wallpaper style
            /// </summary>
            Centered = 0,
            /// <summary>
            /// Streched wallpaper style
            /// </summary>
            Stretched = 1,
            /// <summary>
            /// Tiled wallpaper style
            /// </summary>
            Tiled = 2
        }

        /// <summary>
        /// Set desktop wallpaper
        /// </summary>
        /// <param name="imgPath">Image filepath</param>
        /// <param name="style">Style of wallpaper</param>
        public static void Set(string imgPath, Style style = Style.Stretched)
        {
            try
            {
                //object p = img.Save(tempPath, System.Drawing.Imaging.ImageFormat.Jpeg);

                var key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true);
                if (key == null)
                {
                    return;
                }

                if (style == Style.Stretched)
                {
                    key.SetValue("WallpaperStyle", "2");
                    key.SetValue("TileWallpaper", "0");
                }

                if (style == Style.Centered)
                {
                    key.SetValue("WallpaperStyle", "1");
                    key.SetValue("TileWallpaper", "0");
                }

                if (style == Style.Tiled)
                {
                    key.SetValue("WallpaperStyle", "1");
                    key.SetValue("TileWallpaper", "1");
                }

                SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, imgPath, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
            }
            catch { }
        }
    }
}
