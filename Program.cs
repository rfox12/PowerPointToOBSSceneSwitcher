using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.

namespace PowerPointToOBSSceneSwitcher
{
    class Program
    {
        private static Application ppt = new Microsoft.Office.Interop.PowerPoint.Application();
        private static ObsLocal OBS;
        static async Task Main(string[] args)
        {
            Console.Write("Connecting to PowerPoint...");
            ppt.SlideShowNextSlide += App_SlideShowNextSlide;
            // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa211571(v=office.11)
            // ppt.SlideShowNextClick += App_SlideShowNextClick;  // This event fires on every slide click (even those that don't progress to next slide)
            // ppt.SlideShowNextBuild += App_SlideShowNextBuild;  // Not sure what this even does
            
            ppt.PresentationCloseFinal += App_PresentationCloseFinal;
            // ppt.SlideShowEnd += App_SlideShowEnd;
            Console.WriteLine("connected");

            Console.Write("Connecting to OBS...");
            OBS = new ObsLocal();
            await OBS.Connect();
            Console.WriteLine("connected");

            OBS.GetScenes();

            Console.ReadLine();
        }

        // async static void App_SlideShowNextClick(SlideShowWindow Wn, Effect eff)
        // {
        //     if (Wn != null)
        //     {
        //         Console.WriteLine($"{DateTime.Now.ToString("hh:mm:ss.fff tt")} Click for Slide Number {Wn.View.Slide.SlideNumber}");
        //     }
        // }

        // async static void App_SlideShowNextBuild(SlideShowWindow Wn)
        // {
        //     if (Wn != null)
        //     {
        //         Console.WriteLine($"{DateTime.Now.ToString("hh:mm:ss.fff tt")} Build for Slide Number {Wn.View.Slide.SlideNumber}");
        //     }
        // }

        static void App_PresentationCloseFinal(Presentation p)
        {
            Console.WriteLine("PowerPoint closed!");
            // OBS will disconnect automatically when disposed
            System.Environment.Exit(0);
        }

        // static void App_SlideShowEnd(Presentation p)
        // {
        //     Console.WriteLine("Slide show ended!");
        //     // OBS will disconnect automatically when disposed
        //     System.Environment.Exit(0);
        // }

        async static void App_SlideShowNextSlide(SlideShowWindow Wn)
        {
            if (Wn != null)
            {
                Console.WriteLine($"{DateTime.Now.ToString("hh:mm:ss.fff tt")} Moved to Slide Number {Wn.View.Slide.SlideNumber}");
                //Text starts at Index 2 ¯\_(ツ)_/¯
                var note = String.Empty;
                try { note = Wn.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text; }
                catch { /*no notes*/ }

                bool sceneHandled = false;


                var notereader = new StringReader(note);
                string line;
                while ((line = notereader.ReadLine()) != null)
                {
                    if (line.StartsWith("OBS:"))
                    {
                        line = line.Substring(4).Trim();

                        if (!sceneHandled)
                        {
                            Console.WriteLine($"  Switching to OBS Scene named \"{line}\"");
                            try
                            {
                                sceneHandled = OBS.ChangeScene(line);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"  ERROR: {ex.Message.ToString()}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"  WARNING: Multiple scene definitions found.  I used the first and have ignored \"{line}\"");
                        }
                    }

                    if (line.StartsWith("OBSDEF:"))
                    {
                        OBS.DefaultScene = line.Substring(7).Trim();
                        Console.WriteLine($"  Setting the default OBS Scene to \"{OBS.DefaultScene}\"");
                    }

                    if (line.StartsWith("**START"))
                    {
                        OBS.StartRecording();
                    }

                    if (line.StartsWith("**STOP"))
                    {
                        OBS.StopRecording();
                    }

                    if (!sceneHandled)
                    {
                        OBS.ChangeScene(OBS.DefaultScene);
                        Console.WriteLine($"  Switching to OBS Default Scene named \"{OBS.DefaultScene}\"");
                    }
                }
            }
        }

    }
}