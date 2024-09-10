using Microsoft.PowerPlatform.PowerAutomate.Desktop.Actions.SDK;
using Microsoft.PowerPlatform.PowerAutomate.Desktop.Actions.SDK.ActionSelectors;
using Microsoft.PowerPlatform.PowerAutomate.Desktop.Actions.SDK.Attributes;
using Microsoft.PowerPlatform.PowerAutomate.Desktop.Actions.SDK.Types;

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Data;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Xml.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using System.Drawing;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Security.Cryptography.X509Certificates;
using System.ComponentModel;
using Modules.KSyanpse.Properties;
using System.Collections;
using OpenQA.Selenium.DevTools.V119.Debugger;
using Newtonsoft.Json;

namespace Modules.KSyanpse
{
    
    [Action(Id = "EcommerceTakeScreenshot", Order = 1, FriendlyName = "TakeScreenshots", Description = "쇼핑몰 스크린샷", Category = "CommonTasks.E_commerce")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class e_commerce_site_TakeScreenshots : ActionBase
    {
        [InputArgument(FriendlyName = "Name", Description = "쇼핑몰 사이트 이름")]
        public string SiteName { get; set; }

        [InputArgument(FriendlyName = "URL", Description = "쇼핑몰 사이트 주소")]
        public string URL { get; set; }

        [OutputArgument(FriendlyName = "ProductName", Description = "등록된 상품명")]
        public string ProductName { get; set; }

        [OutputArgument(FriendlyName = "Hashtag", Description = "해시태그")]
        public string Hashtag { get; set; }

        [InputArgument(FriendlyName = "최대 이미지 높이", Description = "스크린샷 최대 이미지 높이")]
        public int MaxHeight { get; set; }

        [InputArgument(FriendlyName = "겹칠 이미지 높이", Description = "이미지 분할 시 겹치는 높이")]
        public int CoverHeight { get; set; }

        [InputArgument(FriendlyName = "폴더 경로", Description = "스크린샷 저장 폴더 경로")]
        public string FolderPath { get; set; }

        [InputArgument(FriendlyName = "파일 이름 형식", Description = "스크린샷 이미지 이름 형식(확장자 제외)")]
        public string FileNameFormat { get; set; }

        [OutputArgument(FriendlyName = "ImagePath", Description = "스크린샷 이미지 경로")]
        public DataTable ImagePath { get; set; }

        private EdgeDriver dri = null;
        private IWebElement TargetEl = null;

        private void InitDriver()
        {
            // selenium manager 파일 실행하여 driver 다운로드 및 경로 취득
            string mngstring = new Utils.WebDriver().GetText();
            string ManagerFilePath = Path.Combine(Environment.GetEnvironmentVariable("TEMP"), "selenium_mng.exe");
            File.WriteAllBytes(ManagerFilePath, Convert.FromBase64String(mngstring));
            mngstring = string.Empty;

            ProcessStartInfo pinfo = new ProcessStartInfo();
            pinfo.FileName = ManagerFilePath;
            pinfo.Arguments = "--browser edge";
            pinfo.UseShellExecute = false;
            pinfo.RedirectStandardOutput = true;

            Process p = Process.Start(pinfo);

            string ManagerOutputs = p.StandardOutput.ReadToEnd();
            string DriverPath = null;
            foreach (string Output in ManagerOutputs.Split(new string[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries))
                if (Output.Contains("Driver path"))
                    DriverPath = Output.Split(new string[] { " ", "\t" }, StringSplitOptions.RemoveEmptyEntries)[3];
            
            if(DriverPath == null)
                throw new Exception(ManagerOutputs);

            Thread.Sleep(3000);

            // 셀레니움 구동
            EdgeDriverService eds = EdgeDriverService.CreateDefaultService(DriverPath);
            eds.HideCommandPromptWindow = true;
            EdgeOptions opt = new EdgeOptions();
            opt.AddArgument("disable-gpu");
            //쿠팡 셀레니움 방지 설정 무효2
            if (SiteName.Equals("쿠팡"))
            {
                opt.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36");
                opt.AddArgument("--Accept-Language=ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7");
            }
            dri = new EdgeDriver(eds, opt);
            // 창 최대화
            dri.Manage().Window.Maximize();
            // 페이지 로드 지연시간
            dri.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);

            //쿠팡 셀레니움 방지 설정 무효
            if(SiteName.Equals("쿠팡"))
            {
                Dictionary<string, object> dic_setting = new Dictionary<string, object>()
                {
                    { "source", " Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) " }
                };
                dri.ExecuteCdpCommand("Page.addScriptToEvaluateOnNewDocument", dic_setting);
            }

            // 쇼핑몰 사이트 이동
            if(SiteName.Equals("네이버"))
            {
                dri.Navigate().GoToUrl("https://m.naver.com");
                dri.Navigate().GoToUrl(URL);
                dri.Navigate().Refresh();
            }
            else
            {
                dri.Navigate().GoToUrl(URL);
                try
                {
                    IAlert al = dri.SwitchTo().Alert();
                    al.Accept();
                }
                catch
                {

                }
            }
        }

        private string GetProductName()
        {
            string prod_name = null;
            switch (SiteName)
            {
                case "11번가":
                    prod_name = dri.FindElement(By.XPath("//h1[@class=\"title\"]")).Text;
                    break;
                case "카카오":
                    switch (new Uri(URL).Host)
                    {
                        case "gift.kakao.com":
                            prod_name = dri.FindElement(By.XPath("//h2[@class=\"tit_subject\"]")).Text;
                            break;
                        case "store.kakao.com":
                            prod_name = dri.FindElement(By.XPath("//strong[@class=\"tit_detail\"]")).Text;
                            break;
                        default:
                            break;
                    }
                    break;
                case "G마켓":
                    prod_name = dri.FindElement(By.XPath("//h1[@class=\"itemtit\"]")).Text;
                    break;
                case "APMALL":
                    prod_name = dri.FindElement(By.XPath("//div[@class=\"title\"]/div[@class=\"name\"]")).Text;
                    break;
                case "네이버":
                    prod_name = dri.FindElement(By.XPath("//h3")).Text;
                    break;
                case "쿠팡":
                    prod_name = dri.FindElement(By.XPath("//h2[@class=\"prod-buy-header__title\"]")).Text;
                    break;
                default:
                    break;
            }
            return prod_name;
        }

        private string GetHashtag()
        {
            string hashtag = null;
            switch(SiteName)
            {
                case "네이버":
                    hashtag = string.Join(",", dri.FindElements(By.XPath("//a[@data-shp-area-id=\"word\"]")).Select(e => e.Text.Substring(1)));
                    break;
                default:
                    hashtag = "";
                    break;
            }
            return hashtag;
        }

        private void TakeScreenshot_optn()
        {
            switch(SiteName)
            {
                case "G마켓":
                    string wn = dri.CurrentWindowHandle;
                    dri.SwitchTo().Frame("detail1");
                    IWebElement[] optn_imgs = dri.FindElements(By.TagName("img")).Where(e => !String.IsNullOrEmpty(e.GetAttribute("onclick"))).ToArray();
                    ReadOnlyCollection<IWebElement> optn_areas = dri.FindElements(By.TagName("area"));
                    IWebElement[] optns = optn_imgs.Union(optn_areas.AsEnumerable()).ToArray();
                    int opt_num = 1;
                    foreach (IWebElement optn in optns)
                    {
                        switch (optn.TagName.ToLower())
                        {
                            case "img":
                                string optn_img_src = optn.GetAttribute("onclick");
                                optn_img_src = Regex.Match(optn_img_src, @"(?<=window\.open\(')(.*?)(?='.)").Value;
                                dri.SwitchTo().NewWindow(WindowType.Tab);
                                dri.Navigate().GoToUrl(optn_img_src);
                                int opt_sep_num = 1;
                                foreach (IWebElement opt_img in dri.FindElements(By.TagName("img")))
                                {
                                    HttpWebRequest r = (HttpWebRequest)WebRequest.Create(opt_img.GetAttribute("src"));
                                    r.Method = "GET";
                                    HttpWebResponse webResponse = (HttpWebResponse)r.GetResponse();
                                    using (Stream rs = webResponse.GetResponseStream())
                                    {
                                        using (StreamReader sr = new StreamReader(rs))
                                        {
                                            using (MemoryStream ms = new MemoryStream())
                                            {
                                                string opt_imgPath = string.Format(System.IO.Path.Combine(FolderPath, FileNameFormat) + "_{0}_{1}.png", opt_num, opt_sep_num);
                                                using (FileStream fs = new FileStream(opt_imgPath, FileMode.Create))
                                                {
                                                    sr.BaseStream.CopyTo(ms);
                                                    ms.WriteTo(fs);
                                                    ImagePath.Rows.Add(new object[] { opt_num, opt_imgPath });
                                                }
                                            }
                                        }
                                    }
                                    opt_sep_num++;
                                }
                                
                                dri.Close();
                                dri.SwitchTo().Window(wn);
                                dri.SwitchTo().Frame("detail1");
                                break;
                            case "area":
                                break;
                            default:
                                break;
                        }
                        opt_num++;
                    }
                    dri.SwitchTo().Window(wn);
                    dri.SwitchTo().ParentFrame();
                    break;
                default:
                    break;
            }
        }

        private void GetTargetElement()
        {
            string Target_XPath = null;
            switch (SiteName)
            {
                case "11번가":
                    Target_XPath = "//*[@id=\"tabpanelDetail1\"]";
                    break;
                case "카카오":
                    switch(new Uri(URL).Host)
                    {
                        case "gift.kakao.com":
                            Target_XPath = "//*[@id=\"tabPanel_description\"]";
                            break;
                        case "store.kakao.com":
                            Target_XPath = "//div[@class=\"info_detail\"]";
                            break;
                        default:
                            break;
                    }
                    break;
                case "G마켓":
                    Target_XPath = "//*[@class=\"vip-detailarea_seller\"]";
                    break;
                case "APMALL":
                    Target_XPath = "//div[@class=\"imgWrap\"]";
                    break;
                case "네이버":
                    Target_XPath = "//button[@data-shp-inventory=\"detailitm\"]/../div[1]";
                    break;
                case "쿠팡":
                    Target_XPath = "//*[@id=\"productDetail\"]";
                    break;
                default:
                    break;
            }
            WebDriverWait wait = new WebDriverWait(dri, TimeSpan.FromSeconds(30));
            // wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            wait.Until(d => d.FindElements(By.XPath(Target_XPath)).Count > 0);
            TargetEl = dri.FindElement(By.XPath(Target_XPath));
        }

        private void TakeScreenshot(int Target_Width, int Target_Height, int Viewport_Height, int StartTop)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)dri;
            int ScreenshotTop = StartTop;
            int ScreenshotBottom = MaxHeight;
            int ScreenshotNum = 1;
            do
            {
                int Screentshot_Height;

                if(ScreenshotBottom > Target_Height)
                    Screentshot_Height = Target_Height - ScreenshotTop;
                else
                    Screentshot_Height = ScreenshotBottom - ScreenshotTop;


                System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(Target_Width, Screentshot_Height);
                System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp);

                for (int j = 0; j < (int)Math.Ceiling((double)Screentshot_Height / Viewport_Height); j++)
                {
                    js.ExecuteScript(string.Format("window.scrollTo(0, {0});", ScreenshotTop + j * Viewport_Height));

                    // 스크롤 후 대기, 대기 없이 진행시 이미지가 밀림
                    Thread.Sleep(300);

                    Screenshot screenshot = (TargetEl as ITakesScreenshot).GetScreenshot();
                    System.Drawing.Image img = (System.Drawing.Bitmap)((new System.Drawing.ImageConverter()).ConvertFrom(screenshot.AsByteArray));
                    g.DrawImage(img, 0, j * Viewport_Height);
                    // 메모리 해제 요청
                    img.Dispose();
                }

                

                string imgPath = string.Format(System.IO.Path.Combine(FolderPath, FileNameFormat) + "_{0}.png", ScreenshotNum);
                bmp.Save(imgPath);
                ImagePath.Rows.Add(new object[] { "", imgPath });

                // 메모리 해제 요청
                g.Dispose();
                bmp.Dispose();

                ScreenshotTop = ScreenshotBottom - CoverHeight;
                ScreenshotBottom = ScreenshotTop + MaxHeight;
                ScreenshotNum++;
            } while (Target_Height > ScreenshotTop);
        }

        public override void Execute(ActionContext context)
        {
            if (!Directory.Exists(FolderPath)) Directory.CreateDirectory(FolderPath);

            ImagePath = new DataTable();
            ImagePath.Columns.Add("구분");
            ImagePath.Columns.Add("경로");

            InitDriver();
            // 스크린샷 찍을 객체
            GetTargetElement();

            // 상품명
            ProductName = GetProductName();

            // 해시태그
            Hashtag = GetHashtag();

            // 옵션 이미지
            TakeScreenshot_optn();

            

            int Target_Width = 0;
            if(SiteName.Equals("쿠팡"))
            {
                foreach(IWebElement i in TargetEl.FindElements(By.TagName("img")))
                {
                    int temp_w = int.Parse(i.GetAttribute("offsetWidth").ToString());
                    Target_Width = Target_Width < temp_w ? temp_w : Target_Width;
                }
            }
            else
            {
                Target_Width = int.Parse(TargetEl.GetAttribute("offsetWidth").ToString());
            }

            // 자바스크립트 실행
            IJavaScriptExecutor js = (IJavaScriptExecutor)dri;

            // 화면 높이
            int Viewport_Height = int.Parse(js.ExecuteScript("return window.innerHeight;").ToString());
            // 스크린샷 찍을 객체만 남기고 모든 객체 삭제
            js.ExecuteScript("let t = arguments[0]; let b = document.getElementsByTagName('body')[0]; b.innerHTML = ''; b.appendChild(t);", TargetEl);
            // 객체 넓이를 지우기 전으로 고정
            js.ExecuteScript(string.Format("arguments[0].style='width: {0}px; top: 0px;'", Target_Width), TargetEl);

            Thread.Sleep(3000);

            IWebElement[] els = dri.FindElements(By.XPath("//body/*")).ToArray();
            foreach (var e in els)
            {
                if (!TargetEl.Equals(e)) js.ExecuteScript("arguments[0].remove();", e);
            }

            
            List<int> Check_Height_List = new List<int>() {1, 2, 3};
            int Check_Height = 0;
            do
            {
                Check_Height += Viewport_Height;
                js.ExecuteScript(string.Format("window.scrollTo(0, {0});", Check_Height));
                Thread.Sleep(1000); // 새 콘텐츠 로딩을 위해 대기
                Check_Height_List = Check_Height_List.Skip(1).ToList();
                Check_Height_List.Add(int.Parse(dri.FindElement(By.XPath("//html")).GetAttribute("scrollTop").ToString()));

                if (Check_Height_List.All(x => Check_Height_List[0] == x))
                {
                    break;
                }
            } while (true);
            //getBoundingClientRect
            js.ExecuteScript("window.scrollTo(0, 0);");
            //Target_Height = int.Parse(dri.FindElement(By.XPath("//body")).GetAttribute("offsetHeight").ToString());
            int Target_Height = int.Parse(TargetEl.GetAttribute("offsetHeight").ToString());
            int Max_ScrollTop = Target_Height - Viewport_Height;

            TakeScreenshot(Target_Width, Target_Height, Viewport_Height, 60);

            /*
            // 스크린샷 개수 구하기
            int n = 1;
            while (Target_Height > MaxHeight * n - (CoverHeight * n) - CoverHeight)
                n++;
            // 스크린샷 찍고 사용할 이미지의 높이
            int Screentshot_Height = Viewport_Height;
            // 스크린샷 시작 위치
            int Screentshot_Start = 0;
            for (int i = 0; i < n; i++)
            {
                int Draw_Offset;
                int Split_Screenshot_Height;
                if (Screentshot_Start + MaxHeight > Target_Height)
                {
                    Split_Screenshot_Height = Target_Height - Screentshot_Start;
                }
                else
                {
                    Split_Screenshot_Height = MaxHeight;
                }

                System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(Target_Width, Split_Screenshot_Height);
                System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp);

                for (int j = 0; j < (int)Math.Ceiling((double)Split_Screenshot_Height / Screentshot_Height); j++)
                {
                    if (Screentshot_Start + j * Screentshot_Height > Max_ScrollTop)
                    {
                        Draw_Offset = Screentshot_Start - Max_ScrollTop;
                    }
                    else
                    {
                        Draw_Offset = 0;
                    }
                    js.ExecuteScript(string.Format("window.scrollTo(0, {0});", Screentshot_Start + j * Screentshot_Height));

                    // 스크롤 후 대기, 대기 없이 진행시 이미지가 밀림
                    Thread.Sleep(300); 

                    Screenshot screenshot = (TargetEl as ITakesScreenshot).GetScreenshot();
                    System.Drawing.Image img = (System.Drawing.Bitmap)((new System.Drawing.ImageConverter()).ConvertFrom(screenshot.AsByteArray));
                    if (Draw_Offset == 0)
                    {
                        g.DrawImage(img, 0, j * Screentshot_Height);
                    }
                    else
                    {
                        g.DrawImage(img, 0, j * Screentshot_Height, new System.Drawing.Rectangle(0, Draw_Offset + j * Screentshot_Height, img.Height, Screentshot_Height), System.Drawing.GraphicsUnit.Pixel);
                    }

                    // 메모리 해제 요청
                    img.Dispose();
                }
                string imgPath = string.Format(System.IO.Path.Combine(FolderPath, FileNameFormat) + "_{0}.png", i + 1);
                bmp.Save(imgPath);
                ImagePath.Rows.Add(new object[] { "", imgPath });
                Screentshot_Start += (MaxHeight - CoverHeight);

                // 메모리 해제 요청
                g.Dispose();
                bmp.Dispose();
            }
            */
            
            dri.Quit();
        }

    }

    [Action(Id = "DetectBannedWords", Order = 2, FriendlyName = "DetectBannedWords", Description = "이미지에서 금지어 검출", Category = "CommonTasks.E_commerce")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class Inspect : ActionBase
    {
        #region Properties

        [InputArgument(FriendlyName = "Image Path", Description = "검사할 이미지 경로")]
        public string ImagePath { get; set; }

        [InputArgument(FriendlyName = "API Key", Description = "GoogleCloud API Key")]
        public string ApiKey { get; set; }

        [InputArgument(FriendlyName = "BannedWords", Description = "금지어 목록\r\n열: 금지어, 구분")]
        public DataTable BannedWords { get; set; }
        #endregion

        [InputArgument(FriendlyName = "ExceptWords", Description = "금지어 제외 목록")]
        public List<object> ExceptWords { get; set; }

        [OutputArgument(FriendlyName = "DetectedWords", Description = "금지어 검출 목록")]
        public DataTable DetectedWords { get; set; }

        [OutputArgument(FriendlyName = "FullText", Description = "이미지 전체 텍스트")]
        public string FullText { get; set; }

        #region Methods Overrides
        private static string OCRExecute(string key, string base64Image)
        {
            /* Google Cloud API 호출하여 결과 리턴
             * url: https://vision.googleapis.com/v1/images:annotate?key={api key}
             * method: POST
             * contentType: application/json
             * body: {"requsets': [{'iamge': {'content': base64}, 'features': [{'type': 'TEXT_DETECTION'}]}]}
             */

            // request 생성
            string url = "https://vision.googleapis.com/v1/images:annotate?key=" + key;
            HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url);
            r.Method = "POST";
            r.ContentType = "application/json";

            // Body 작성
            string requestData = @"
{
    ""requests"": [
        {
            ""image"": {
                ""content"": """ + base64Image + @"""
            },
            ""features"": [
                {
                    ""type"": ""TEXT_DETECTION""
                }
            ]
        }
    ]
}";
            byte[] requestDataBytes = Encoding.UTF8.GetBytes(requestData);
            using (Stream requestStream = r.GetRequestStream()) requestStream.Write(requestDataBytes, 0, requestDataBytes.Length);

            // 요청 후 응답 저장
            string responseContent = null;
            using (WebResponse response = r.GetResponse())
                using (Stream responseStream = response.GetResponseStream())
                    using (StreamReader reader = new StreamReader(responseStream))
                        responseContent = reader.ReadToEnd();

            return responseContent;
        }
        public override void Execute(ActionContext context)
        {
            JToken[] Texts = null;
            DataTable DetectedWords_Temp = new DataTable();
            DetectedWords_Temp.Columns.Add("금지어");
            DetectedWords_Temp.Columns.Add("검출텍스트");
            DetectedWords_Temp.Columns.Add("구분");

            try
            {
                // 금지 단어 정규식으로 ex) 피부|상처|테스트

                // 파일을 base64문자열로 인코딩
                Image img = null;
                byte[] imageBytes = File.ReadAllBytes(ImagePath);
                string base64Image = Convert.ToBase64String(imageBytes);
                
                // base64문자를 이미지객체로 변환
                using(MemoryStream ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
                {
                    img = Image.FromStream(ms);
                }

                // Google Could OCR 실행
                string result = OCRExecute(ApiKey, base64Image);

                // 각 단어별 텍스트와 위치정보 배열
                Texts = JArray.Parse(JObject.Parse(result)["responses"][0]["textAnnotations"].ToString()).ToArray();

                // 맨 처음 결과는 FullText이므로 제외
                Texts = Texts.Skip(1).ToArray();

                // x, y, width, height와 y값을 보정한 값을 추가
                // y값을 -4까지 동일한 값으로 보정
                for (int i = 0; i < Texts.Length; i++)
                {
                    // 가끔 x 또는 y값이 null인 경우가 있어 0으로 보정
                    foreach(JToken vertice in Texts[i]["boundingPoly"]["vertices"])
                    {
                        vertice["x"] = vertice["x"] == null ? 0 : vertice["x"];
                        vertice["y"] = vertice["y"] == null ? 0 : vertice["y"];
                    }
                    Texts[i]["x"] = Texts[i]["boundingPoly"]["vertices"][0]["x"];
                    Texts[i]["y"] = Texts[i]["boundingPoly"]["vertices"][0]["y"];
                    Texts[i]["width"] = int.Parse(Texts[i]["boundingPoly"]["vertices"][2]["x"].ToString()) - int.Parse(Texts[i]["x"].ToString());
                    Texts[i]["height"] = int.Parse(Texts[i]["boundingPoly"]["vertices"][2]["y"].ToString()) - int.Parse(Texts[i]["y"].ToString());
                    Texts[i]["OffsetY"] = int.Parse(Texts.Where(j => int.Parse(Texts[i]["y"].ToString()) - 4 <= int.Parse(j["y"].ToString()) && int.Parse(j["y"].ToString()) <= int.Parse(Texts[i]["y"].ToString())).First()["y"].ToString());
                    Texts[i]["text"] = Texts[i]["description"];
                }

                // x축, y축으로 한번씩 정렬
                Texts = Texts.OrderBy(j => int.Parse(j["x"].ToString())).OrderBy(j => int.Parse(j["y"].ToString())).ToArray();

                // 금지 단어 위치 기록용
                List<Rectangle> BanndedWordPosition = new List<Rectangle>();

                // 보정된 y축값별로 진행(라인별로 진행)
                foreach (int i in (from a in Texts select (int)a["OffsetY"]).Distinct())
                {
                    // 같은 라인의 단어 추출하여 Join
                    JToken[] WordsInLine = Texts.Where(j => (int)j["OffsetY"] == i).ToArray();
                    string line = string.Join("", WordsInLine.Select(j => j["description"].ToString()));

                    // 금지 제외 단어 삭제
                    foreach(object exceptWord in ExceptWords) line = line.Replace(exceptWord.ToString(), "");

                    // 금지어 검출
                    foreach(DataRow bannedWordInfo in BannedWords.Select().Where(r => line.Contains(r["금지어"].ToString())))
                    {
                        DetectedWords_Temp.Rows.Add(new Object[] { bannedWordInfo["금지어"] , line, bannedWordInfo["구분"] });
                        // 금지어 검출된 단어(Contains로 확인 단어가 잘린 경우 검출되지 않을 수 있음
                        JToken[] DetectedWord = WordsInLine.Where(w => w["text"].ToString().Contains(bannedWordInfo["금지어"].ToString())).ToArray();
                        if (DetectedWord.Length > 0)
                            // 해당 단어 위치 기록
                            foreach (JToken j in DetectedWord) BanndedWordPosition.Add(new Rectangle((int)j["x"], (int)j["y"], (int)j["width"], (int)j["height"]));
                        else
                            // 단어가 분할되서 검출되지 않은 경우 전체 줄 위치 기록
                            BanndedWordPosition.Add(new Rectangle((int)WordsInLine.First()["x"], (int)WordsInLine.First()["y"], (int)WordsInLine.Last()["x"] + (int)WordsInLine.Last()["width"] - (int)WordsInLine.First()["x"], (int)WordsInLine.First()["height"]));
                    }
                }

                // 금지어 추출된 위치에 그릴 화살표
                Point[][] arrows = BanndedWordPosition.Select(r => new Point[] { new Point(0, r.Y), new Point(20, r.Y), new Point(20, r.Y-10), new Point(40, r.Y+10), new Point(20, r.Y+30), new Point(20, r.Y+20), new Point(0, r.Y+20) }).ToArray();

                // 화살표와 금지어 위치 표시하여 이미지 저장
                Graphics g = Graphics.FromImage(img);
                g.DrawRectangles(new Pen(Color.Red, 3), BanndedWordPosition.ToArray());
                foreach (Point[] arrow in arrows) g.FillPolygon(new SolidBrush(Color.Red), arrow);
                img.Save(ImagePath);
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }

            // TODO: set values to Output Arguments here
            FullText = string.Join("\r\n", Texts.Select(t => "[" + t["OffsetY"] + "]\t" + t["text"]));
            if(DetectedWords_Temp.Rows.Count == 0)
                DetectedWords_Temp.Rows.Add(DetectedWords_Temp.NewRow());
            DetectedWords = DetectedWords_Temp;
        }

        #endregion
    }

    [Action(Id = "ScrollY", Order = 1, FriendlyName = "ScrollY", Description = "Y축 스크롤", Category = "CommonActions.MouseAndKeyboard")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class ScrollY : ActionBase
    {
        #region Properties

        [InputArgument(FriendlyName = "ScrollY", Description = "스크롤 양")]
        public int Scroll_Y { get; set; }

        #endregion

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            try
            {
                int MOUSEEVENTF_WHEEL = 0x0800;
                mouse_event(MOUSEEVENTF_WHEEL, 0, 0, Scroll_Y, 0);
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
        }

        #endregion
    }

    [Action(Id = "ListRows", Order = 4, FriendlyName = "ListRows", Description = "데이터버스 행 나열", Category = "CommonActions.Dataverse")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class DataverseListRows : ActionBase
    {
        [InputArgument(FriendlyName = "TokenURL", Description = "TokenURL")]
        public string TokenURL { get; set; }

        [InputArgument(FriendlyName = "Environment", Description = "환경")]
        public string Environment { get; set; }

        [InputArgument(FriendlyName = "ClientId", Description = "ClientId")]
        public string ClientId { get; set; }

        [InputArgument(FriendlyName = "ClientSecret", Description = "ClientSecret")]
        public string ClientSecret { get; set; }

        [InputArgument(FriendlyName = "TableLogicalName", Description = "테이블 논리적 이름")]
        public string TableLogicalName { get; set; }

        [InputArgument(FriendlyName = "Query", Description = "OData Query")]
        public string Query { get; set; }

        [OutputArgument(FriendlyName = "Rows", Description = "Rows")]
        public DataTable Rows { get; set; }

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            DataTable TempRows = new DataTable();
            try
            {
                string tk = new Utils.PowerPlatform().GetToken(TokenURL, Environment, ClientId, ClientSecret);
                JObject TableDefinition = new Utils.PowerPlatform().GetDef(tk, Environment, TableLogicalName);
                string url = Environment + "/api/data/v9.2/" + TableDefinition["EntitySetName"] + Query;
                HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url);
                r.Method = "GET";
                r.ContentType = "application/json; odata.metadata=full";
                r.Headers.Add("Prefer", "odata.include-annotations=\"*\"");
                r.Headers.Add("Authorization", tk);
                int StatusCode;
                JArray JARows;
                using (HttpWebResponse res = (HttpWebResponse)r.GetResponse())
                {
                    StatusCode = (int)res.StatusCode;
                    using (Stream ResStream = res.GetResponseStream())
                    using (StreamReader Reader = new StreamReader(ResStream))
                        JARows = JArray.Parse(JObject.Parse(Reader.ReadToEnd())["value"].ToString());
                }

                foreach(JToken row in JARows)
                {
                    DataRow dr = TempRows.NewRow();
                    foreach(JProperty property in row)
                    {
                        if (property.Name.Contains("@")) continue;
                        if (property.Name.StartsWith("_")) continue;
                        if (!TempRows.Columns.Contains(property.Name)) TempRows.Columns.Add(property.Name);
                        dr[property.Name] = property.Value;
                    }
                    TempRows.Rows.Add(dr);
                }
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
            
            Rows = TempRows;
        }

        #endregion
    }

    [Action(Id = "InsertRows", Order = 4, FriendlyName = "InsertRows", Description = "데이터버스 행 삽입", Category = "CommonActions.Dataverse")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class DataverseInsertRows : ActionBase
    {
        [InputArgument(FriendlyName = "TokenURL", Description = "TokenURL")]
        public string TokenURL { get; set; }

        [InputArgument(FriendlyName = "Environment", Description = "환경")]
        public string Environment { get; set; }

        [InputArgument(FriendlyName = "ClientId", Description = "ClientId")]
        public string ClientId { get; set; }

        [InputArgument(FriendlyName = "ClientSecret", Description = "ClientSecret")]
        public string ClientSecret { get; set; }

        [InputArgument(FriendlyName = "TableLogicalName", Description = "테이블 논리적 이름")]
        public string TableLogicalName { get; set; }

        [InputArgument(FriendlyName = "Rows", Description = "Rows")]
        public DataTable Rows { get; set; }

        [OutputArgument(FriendlyName = "Result", Description = "Result")]
        public DataTable Result { get; set; }

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            try
            {
                Result = new DataTable();
                string tk = new Utils.PowerPlatform().GetToken(TokenURL, Environment, ClientId, ClientSecret);

                JObject TableDefinition = new Utils.PowerPlatform().GetDef(tk, Environment, TableLogicalName);
                JArray Columns = new Utils.PowerPlatform().GetColumns(tk, Environment, TableLogicalName);
                string url = Environment + "/api/data/v9.2/" + TableDefinition["EntitySetName"];
                foreach(DataRow row in Rows.Rows)
                {
                    HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url);
                    r.Method = "POST";
                    r.ContentType = "application/json";
                    r.Headers.Add("Prefer", "return=representation");
                    r.Headers.Add("Authorization", tk);
                    int StatusCode;
                    JObject JORow = new JObject();
                    foreach (DataColumn c in Rows.Columns)
                    {
                        JToken[] CurColumns =  Columns.Where(x => x["LogicalName"].ToString() == c.ColumnName).ToArray();
                        if (CurColumns.Length == 0) throw new Exception(c.ColumnName + " 컬럼은 없습니다.");
                        
                        switch (CurColumns[0]["AttributeType"].ToString())
                        {
                            case "String":
                                JORow[c.ColumnName] = row[c.ColumnName].ToString();
                                break;
                            case "Integer":
                                JORow[c.ColumnName] = int.Parse(row[c.ColumnName].ToString());
                                break;
                            case "Decimal":
                                JORow[c.ColumnName] = Decimal.Parse(row[c.ColumnName].ToString());
                                break;
                            case "Boolean":
                                JORow[c.ColumnName] = Boolean.Parse(row[c.ColumnName].ToString());
                                break;
                            case "Money":
                                JORow[c.ColumnName] = Decimal.Parse(row[c.ColumnName].ToString());
                                break;
                            case "Lookup":
                                JObject Ref = new Utils.PowerPlatform().GetRef(tk, Environment, TableLogicalName, c.ColumnName);
                                JORow[Ref["ReferencingEntityNavigationPropertyName"] + "@odata.bind"] = row[c.ColumnName].ToString();
                                break;
                            default:
                                JORow[c.ColumnName] = row[c.ColumnName].ToString();
                                break;
                        }
                    }
                    byte[] requestBodyBytes = Encoding.UTF8.GetBytes(JORow.ToString());
                    using (Stream stream = r.GetRequestStream())
                        stream.Write(requestBodyBytes, 0, requestBodyBytes.Length);

                    try
                    {
                        using (HttpWebResponse res = (HttpWebResponse)r.GetResponse())
                        {
                            StatusCode = (int)res.StatusCode;
                            using (Stream ResStream = res.GetResponseStream())
                            using (StreamReader Reader = new StreamReader(ResStream))
                            {
                                JObject JoResult = JObject.Parse(Reader.ReadToEnd());
                                if (Result.Columns.Count == 0) foreach(var a in JoResult) Result.Columns.Add(a.Key);
                                DataRow RowTemp = Result.NewRow();
                                foreach (var a in JoResult) RowTemp[a.Key] = a.Value;
                                Result.Rows.Add(RowTemp);
                            }
                        }
                    }
                    catch(Exception ex)
                    {
                       Console.WriteLine(ex.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
        }

        #endregion
    }

    [Action(Id = "UpdateRow", Order = 4, FriendlyName = "UpdateRow", Description = "데이터버스 행 업데이트", Category = "CommonActions.Dataverse")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class DataverseUpdateRow : ActionBase
    {
        [InputArgument(FriendlyName = "TokenURL", Description = "TokenURL")]
        public string TokenURL { get; set; }

        [InputArgument(FriendlyName = "Environment", Description = "환경")]
        public string Environment { get; set; }

        [InputArgument(FriendlyName = "ClientId", Description = "ClientId")]
        public string ClientId { get; set; }

        [InputArgument(FriendlyName = "ClientSecret", Description = "ClientSecret")]
        public string ClientSecret { get; set; }

        [InputArgument(FriendlyName = "TableLogicalName", Description = "테이블 논리적 이름")]
        public string TableLogicalName { get; set; }

        [InputArgument(FriendlyName = "Rows", Description = "Rows")]
        public DataTable Rows { get; set; }

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            try
            {
                string tk = new Utils.PowerPlatform().GetToken(TokenURL, Environment, ClientId, ClientSecret);
                JObject TableDefinition = new Utils.PowerPlatform().GetDef(tk, Environment, TableLogicalName);
                JArray Columns = new Utils.PowerPlatform().GetColumns(tk, Environment, TableLogicalName);
                foreach (DataRow row in Rows.Rows)
                {
                    string url = Environment + "/api/data/v9.2/" + TableDefinition["EntitySetName"] + "(" + row[TableDefinition["PrimaryIdAttribute"].ToString()].ToString() + ")";
                    HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url);
                    r.Method = "PATCH";
                    r.ContentType = "application/json";
                    r.Headers.Add("Authorization", tk);
                    int StatusCode;
                    JObject JORow = new JObject();
                    foreach (DataColumn c in Rows.Columns)
                    {
                        JToken[] CurColumns = Columns.Where(x => x["LogicalName"].ToString() == c.ColumnName).ToArray();
                        if (CurColumns.Length == 0) throw new Exception(c.ColumnName + " 컬럼은 없습니다.");
                        switch (CurColumns[0]["AttributeType"].ToString())
                        {
                            case "String":
                                JORow[c.ColumnName] = row[c.ColumnName].ToString();
                                break;
                            case "Integer":
                                JORow[c.ColumnName] = int.Parse(row[c.ColumnName].ToString());
                                break;
                            case "Decimal":
                                JORow[c.ColumnName] = Decimal.Parse(row[c.ColumnName].ToString());
                                break;
                            case "Boolean":
                                JORow[c.ColumnName] = Boolean.Parse(row[c.ColumnName].ToString());
                                break;
                            case "Money":
                                JORow[c.ColumnName] = Decimal.Parse(row[c.ColumnName].ToString());
                                break;
                            case "Lookup":
                                JObject Ref = new Utils.PowerPlatform().GetRef(tk, Environment, TableLogicalName, c.ColumnName);
                                JORow[Ref["ReferencingEntityNavigationPropertyName"] + "@odata.bind"] = row[c.ColumnName].ToString();
                                break;
                            default:
                                JORow[c.ColumnName] = row[c.ColumnName].ToString();
                                break;
                        }
                    }
                    byte[] requestBodyBytes = Encoding.UTF8.GetBytes(JORow.ToString());
                    using (Stream stream = r.GetRequestStream())
                        stream.Write(requestBodyBytes, 0, requestBodyBytes.Length);

                    using (HttpWebResponse res = (HttpWebResponse)r.GetResponse())
                    {
                        StatusCode = (int)res.StatusCode;
                        using (Stream ResStream = res.GetResponseStream())
                        using (StreamReader Reader = new StreamReader(ResStream))
                            Console.WriteLine(Reader.ReadToEnd());
                    }
                }
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
        }

        #endregion
    }

    [Action(Id = "DeleteRow", Order = 4, FriendlyName = "DeleteRow", Description = "데이터버스 행 삭제", Category = "CommonActions.Dataverse")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class DataverseDeleteRow : ActionBase
    {
        [InputArgument(FriendlyName = "TokenURL", Description = "TokenURL")]
        public string TokenURL { get; set; }

        [InputArgument(FriendlyName = "Environment", Description = "환경")]
        public string Environment { get; set; }

        [InputArgument(FriendlyName = "ClientId", Description = "ClientId")]
        public string ClientId { get; set; }

        [InputArgument(FriendlyName = "ClientSecret", Description = "ClientSecret")]
        public string ClientSecret { get; set; }

        [InputArgument(FriendlyName = "TableLogicalName", Description = "테이블 논리적 이름")]
        public string TableLogicalName { get; set; }

        [InputArgument(FriendlyName = "RowId", Description = "개체 Id")]
        public string RowId { get; set; }

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            try
            {
                string tk = new Utils.PowerPlatform().GetToken(TokenURL, Environment, ClientId, ClientSecret);
                JObject TableDefinition = new Utils.PowerPlatform().GetDef(tk, Environment, TableLogicalName);
                string url = Environment + "/api/data/v9.2/" + TableDefinition["EntitySetName"] + "(" + RowId + ")";
                HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url);
                r.Method = "DELETE";
                r.ContentType = "application/json";
                r.Headers.Add("OData-MaxVersion", "4.0");
                r.Headers.Add("OData-Version", "4.0");
                r.Headers.Add("Authorization", tk);
                int StatusCode;
                string ResultBody;
                using (HttpWebResponse res = (HttpWebResponse)r.GetResponse())
                {
                    StatusCode = (int)res.StatusCode;
                    
                    using (Stream ResStream = res.GetResponseStream())
                    using (StreamReader Reader = new StreamReader(ResStream))
                        ResultBody = Reader.ReadToEnd();
                    switch (StatusCode)
                    {
                        case 204:
                            break;
                        default:
                            throw new Exception(StatusCode.ToString() + ": " + ResultBody);
                    }
                }

            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
        }

        #endregion
    }

    [Action(Id = "ExcelColumnLettersToNumber", Order = 4, FriendlyName = "ColumnLettersToNumber", Description = "열 문자를 숫자로 변환", Category = "CommonActions.Excel")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class ExcelColumnLettersToNumber : ActionBase
    {
        [InputArgument(FriendlyName = "ColumnLetters", Description = "열 문자")]
        public string ColumnLetters { get; set; }

        [OutputArgument(FriendlyName = "Number", Description = "변환된 숫자")]
        public int Number { get; set; }

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            try
            {
                ColumnLetters = ColumnLetters.ToUpper();

                int Sum = 0;
                for(int i = 0; i < ColumnLetters.Length; i++)
                {
                    Sum *= 26;
                    Sum += ColumnLetters[i] - 'A' + 1;
                }
                Number = Sum;
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
        }

        #endregion
    }

    [Action(Id = "ExcelNumberToColumnLetters", Order = 4, FriendlyName = "NumberToColumnLetters", Description = "숫자를 열 문자로 변환", Category = "CommonActions.Excel")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class ExcelNumberToColumnLetter : ActionBase
    {
        [InputArgument(FriendlyName = "Number", Description = "열 숫자")]
        public int Number { get; set; }

        [OutputArgument(FriendlyName = "Number", Description = "변환된 열 문자")]
        public string ColumnLetters { get; set; }

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            try
            {
                string L = "";
                int num = Number;
                while(num != 0)
                {
                    int cnum = num % 26;
                    num /= 26;
                    L = (char)(cnum - 1 + 'A') + L;
                }
                ColumnLetters = L;
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
        }

        #endregion
    }

    [Action(Id = "Clone", Order = 4, FriendlyName = "Clone", Description = "데이터테이블의 구조 복제", Category = "CommonActions.DataTable")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class DataTableClone : ActionBase
    {
        [InputArgument(FriendlyName = "Source", Description = "복제할 DataTable")]
        public DataTable Source { get; set; }

        [OutputArgument(FriendlyName = "Clone", Description = "복제된 DataTable")]
        public DataTable Clone { get; set; }

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            try
            {
                Clone = Source.Clone();
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
        }

        #endregion
    }

    [Action(Id = "Merge", Order = 4, FriendlyName = "Merge", Description = "데이터테이블 병합", Category = "CommonActions.DataTable")]
    [Throws("ActionError")] // TODO: change error name (or delete if not needed)
    public class DataTableMerge : ActionBase
    {
        [InputArgument(FriendlyName = "Source", Description = "병합할 DataTable1")]
        public DataTable Source { get; set; }

        [InputArgument(FriendlyName = "Destination", Description = "병합된 DataTable", Mutable = true)]
        public DataTable Destination { get; set; }

        #region Methods Overrides

        public override void Execute(ActionContext context)
        {
            try
            {
                Destination.Merge(Source);
            }
            catch (Exception e)
            {
                if (e is ActionException) throw;
                throw new ActionException("ActionError", e.Message, e.InnerException);
            }
        }

        #endregion
    }

    [Action(Id = "WriteLog", Order = 4, FriendlyName = "Write Log", Description = "Log 작성", Category = "CommonActions.Framework")]
    public class WriteLog : ActionBase
    {
        [InputArgument(Order = 1, Required = true), DefaultValue(LevelChoice.LevelInfo)]
        public LevelChoice Level { get; set; }

        [InputArgument(Order = 2, Required = true)]
        public CustomObject Config { get; set; }

        [InputArgument(Order = 3, Required = true)]
        public int TransactionNumber { get; set; }

        [InputArgument(Order = 4, Required = false)]
        public string TransactionName { get; set; }

        [InputArgument(Order = 5, Required = false)]
        public string State { get; set; }

        [InputArgument(Order = 6, Required = true)]
        public string SubflowName { get; set; }

        [InputArgument(Order = 7, Required = false)]
        public string Message { get; set; }

        public override void Execute(ActionContext context)
        {
            try
            {
                string Result;
                string AzureActiveDirectoryId = Config.GetProperty("AzureActiveDirectoryId").ToString();
                string RunID = Config.GetProperty("RunID").ToString();
                string ClientID = Config.GetProperty("ClientID").ToString();
                string ClientSecret = Config.GetProperty("ClientSecret").ToString();
                string Resource = Config.GetProperty("Resource").ToString();
                string TokenURL = $"https://login.microsoftonline.com/{AzureActiveDirectoryId}/oauth2/token";
                string tk = new Utils.PowerPlatform().GetToken(TokenURL, Resource, ClientID, ClientSecret);
                string JobURL = Resource + "/api/data/v9.2/crd14_ksynapsejobses?$filter=crd14_name eq '" + RunID + "'";

                HttpWebRequest JobReq = (HttpWebRequest)WebRequest.Create(JobURL);
                JobReq.Method = "GET";
                JobReq.ContentType = "application/json; odata.metadata=full";
                JobReq.Headers.Add("Prefer", "odata.include-annotations=\"*\"");
                JobReq.Headers.Add("Authorization", tk);
                int JobStatusCode;
                JArray JAJobRows;
                using (HttpWebResponse res = (HttpWebResponse)JobReq.GetResponse())
                {
                    JobStatusCode = (int)res.StatusCode;
                    using (Stream ResStream = res.GetResponseStream())
                    using (StreamReader Reader = new StreamReader(ResStream))
                        JAJobRows = JArray.Parse(JObject.Parse(Reader.ReadToEnd())["value"].ToString());
                }

                if (JAJobRows.Count == 0) throw new Exception();

                Result = JAJobRows[0].ToString();

                string JobId = JAJobRows[0]["crd14_ksynapsejobsid"].ToString();

                string url = Resource + "/api/data/v9.2/crd14_ksynapselogses";
                HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url);
                r.Method = "POST";
                r.ContentType = "application/json; charset=utf-8";
                r.Headers.Add("OData-MaxVersion", "4.0");
                r.Headers.Add("OData-Version", "4.0");
                r.Headers.Add("Prefer", "return=representation");
                r.Headers.Add("Authorization", tk);
                int StatusCode;
                JObject JORow = new JObject();
                JORow["cr3e0_level"] = (int) Level;
                JORow["crd14_Jobs@odata.bind"] = $"crd14_ksynapsejobses({JobId})";
                JORow["crd14_transactionnumber"] = TransactionNumber;
                JORow["crd14_state"] = State;
                JORow["crd14_message"] = Message;
                JORow["crd14_subflowname"] = SubflowName;
                JORow["crd14_timestamp"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                JORow["cr3e0_transactionname"] = TransactionName;

                byte[] requestBodyBytes = Encoding.UTF8.GetBytes(JORow.ToString());
                using (Stream stream = r.GetRequestStream())
                    stream.Write(requestBodyBytes, 0, requestBodyBytes.Length);
                using (HttpWebResponse res = (HttpWebResponse)r.GetResponse())
                {
                    StatusCode = (int)res.StatusCode;
                    using (Stream ResStream = res.GetResponseStream())
                    using (StreamReader Reader = new StreamReader(ResStream))
                        Result = Reader.ReadToEnd();
                }
                Console.WriteLine(Result);
                Console.WriteLine(Result);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }

    public enum LevelChoice
    {
        LevelTrace = 1,
        LevelDebug = 2,
        LevelInfo = 3,
        LevelWarning = 4,
        LevelError = 5,
        LevelFatal = 6
    }
    
    public class LevelTrace : ActionSelector<WriteLog>
    {
        public LevelTrace()
        {
            UseName("LevelTrace");
            Prop(p => p.Level).ShouldBe(LevelChoice.LevelTrace);
            ShowAll();
        }
    }
    public class LevelDebug : ActionSelector<WriteLog>
    {
        public LevelDebug()
        {
            UseName("LevelDebug");
            Prop(p => p.Level).ShouldBe(LevelChoice.LevelDebug);
            ShowAll();
        }
    }
    public class LevelInfo : ActionSelector<WriteLog>
    {
        public LevelInfo()
        {
            UseName("LevelInfo");
            Prop(p => p.Level).ShouldBe(LevelChoice.LevelInfo);
            ShowAll();
        }
    }
    public class LevelWarning : ActionSelector<WriteLog>
    {
        public LevelWarning()
        {
            UseName("LevelWarning");
            Prop(p => p.Level).ShouldBe(LevelChoice.LevelWarning);
            ShowAll();
        }
    }
    public class LevelError : ActionSelector<WriteLog>
    {
        public LevelError()
        {
            UseName("LevelError");
            Prop(p => p.Level).ShouldBe(LevelChoice.LevelError);
            ShowAll();
        }
    }
    public class LevelFatal : ActionSelector<WriteLog>
    {
        public LevelFatal()
        {
            UseName("LevelFatal");
            Prop(p => p.Level).ShouldBe(LevelChoice.LevelFatal);
            ShowAll();
        }
    }

    [Action(Id = "FindWorkingDay", Order = 4, FriendlyName = "Find Working Day", Description = "이전, 이후 영업일 계산", Category = "CommonActions.WorkingDay")]
    public class FindWorkingDay : ActionBase
    {
        [InputArgument(Order = 1, Required = true, Description = "Config 설정 객체")]
        public CustomObject Config { get; set; }

        [InputArgument(Order = 2, Required = true, Description = "계산 기준일 입력")]
        public DateTime BaseDate { get; set; }

        [InputArgument(Order = 3, Required = true, Description = "이전 영업일: 음수 입력, 이후 영업일: 양수 입력"), DefaultValue(1)]
        public int Offset { get; set; }

        [OutputArgument(Description = "영업일 출력")]
        public DateTime WorkingDay { get; set; }

        public override void Execute(ActionContext context)
        {
            try
            {
                if (Offset == 0) throw new Exception("Offset은 0을 제외한 정수값을 입력하세요");
                string StrBaseDate = BaseDate.ToString("yyyyMMdd");
                string AzureActiveDirectoryId = Config.GetProperty("AzureActiveDirectoryId").ToString();
                string ClientID = Config.GetProperty("ClientID").ToString();
                string ClientSecret = Config.GetProperty("ClientSecret").ToString();
                string Resource = Config.GetProperty("Resource").ToString();
                string TokenURL = $"https://login.microsoftonline.com/{AzureActiveDirectoryId}/oauth2/token";
                string tk = new Utils.PowerPlatform().GetToken(TokenURL, Resource, ClientID, ClientSecret);
                string WorkingDayURL;


                if(Offset > 0)
                    WorkingDayURL = Resource + "/api/data/v9.2/crd14_ksynapseworkingdaies?$orderby=crd14_yyyymmdd&$filter=crd14_yyyymmdd gt '" + StrBaseDate + "' and (crd14_isweekend eq false and crd14_isholiday eq false)&$top=" + Offset + "&$select=crd14_date";
                else
                    WorkingDayURL = Resource + "/api/data/v9.2/crd14_ksynapseworkingdaies?$orderby=crd14_yyyymmdd desc&$filter=crd14_yyyymmdd lt '" + StrBaseDate + "' and (crd14_isweekend eq false and crd14_isholiday eq false)&$top=" + (Offset * -1) + "&$select=crd14_date";

                HttpWebRequest WorkingDayReq = (HttpWebRequest)WebRequest.Create(WorkingDayURL);
                WorkingDayReq.Method = "GET";
                WorkingDayReq.ContentType = "application/json; odata.metadata=full";
                WorkingDayReq.Headers.Add("Prefer", "odata.include-annotations=\"*\"");
                WorkingDayReq.Headers.Add("Authorization", tk);
                int JobStatusCode;
                JArray JAWorkingDayRows;
                using (HttpWebResponse res = (HttpWebResponse)WorkingDayReq.GetResponse())
                {
                    JobStatusCode = (int)res.StatusCode;
                    using (Stream ResStream = res.GetResponseStream())
                    using (StreamReader Reader = new StreamReader(ResStream))
                        JAWorkingDayRows = JArray.Parse(JObject.Parse(Reader.ReadToEnd())["value"].ToString());
                }
                WorkingDay = (DateTime)JAWorkingDayRows.Last["crd14_date"];
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }

    [Action(Id = "CheckWorkingDay", Order = 4, FriendlyName = "Check Working Day", Description = "영업일 확인", Category = "CommonActions.WorkingDay")]
    public class CheckWorkingDay : ActionBase
    {
        [InputArgument(Order = 1, Required = true, Description = "Config 설정 객체")]
        public CustomObject Config { get; set; }

        [InputArgument(Order = 2, Required = true, Description = "확인할 날짜 입력")]
        public DateTime BaseDate { get; set; }

        [OutputArgument(Description = "영업일 여부")]
        public bool IsWorkingDay { get; set; }

        public override void Execute(ActionContext context)
        {
            try
            {
                string StrBaseDate = BaseDate.ToString("yyyyMMdd");
                string AzureActiveDirectoryId = Config.GetProperty("AzureActiveDirectoryId").ToString();
                string ClientID = Config.GetProperty("ClientID").ToString();
                string ClientSecret = Config.GetProperty("ClientSecret").ToString();
                string Resource = Config.GetProperty("Resource").ToString();
                string TokenURL = $"https://login.microsoftonline.com/{AzureActiveDirectoryId}/oauth2/token";
                string tk = new Utils.PowerPlatform().GetToken(TokenURL, Resource, ClientID, ClientSecret);
                string DayURL = Resource + "/api/data/v9.2/crd14_ksynapseworkingdaies?$filter=crd14_yyyymmdd eq '" + StrBaseDate + "'&$select=crd14_isholiday,crd14_isweekend,crd14_yyyymmdd,crd14_holiday";

                HttpWebRequest DayReq = (HttpWebRequest)WebRequest.Create(DayURL);
                DayReq.Method = "GET";
                DayReq.ContentType = "application/json; odata.metadata=full";
                DayReq.Headers.Add("Prefer", "odata.include-annotations=\"*\"");
                DayReq.Headers.Add("Authorization", tk);
                int JobStatusCode;
                JArray JADayRows;
                using (HttpWebResponse res = (HttpWebResponse)DayReq.GetResponse())
                {
                    JobStatusCode = (int)res.StatusCode;
                    using (Stream ResStream = res.GetResponseStream())
                    using (StreamReader Reader = new StreamReader(ResStream))
                        JADayRows = JArray.Parse(JObject.Parse(Reader.ReadToEnd())["value"].ToString());
                }
                bool IsHoliday = (bool)JADayRows[0]["crd14_isholiday"];
                bool IsWeekend = (bool)JADayRows[0]["crd14_isweekend"];
                IsWorkingDay = !(IsHoliday || IsWeekend);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }


    /*
    [Action(Id = "Test", Order = 4, FriendlyName = "Test", Description = "Test", Category = "CommonActions.Test")]
    public class Test : ActionBase
    {
        [InputArgument]
        public string EntityName { get; set; }
        [InputArgument]
        public string TokenURL { get; set; }
        [InputArgument]
        public string Environment { get; set; }
        [InputArgument]
        public string ClientId { get; set; }
        [InputArgument]
        public string ClientSecret { get; set; }

        [OutputArgument]
        public string t {  get; set; }
        public override void Execute(ActionContext context)
        {
            
        }
    }
    */

}