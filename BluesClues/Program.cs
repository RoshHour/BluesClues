using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace BluesClues {
    class Program {
        static void Main(string[] args) {
            Console.Title = "Blue's Clues";

            while (true) {
                Console.Clear();
                Console.WriteLine("Enter text to search for:");
                string search = Console.ReadLine().ToLower();

                foreach (string path in GetAllPowerpoints()) {
                    var app = new ApplicationClass();
                    var ppt = app.Presentations.Open(path, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    for (int i = 0; i < ppt.Slides.Count; i++) {
                        var slide = ppt.Slides[i + 1];

                        foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes) {
                            if (shape.HasTextFrame == MsoTriState.msoTrue) {

                                var textFrame = shape.TextFrame;
                                if (textFrame.HasText == MsoTriState.msoTrue) {
                                    var textRange = textFrame.TextRange.Text.ToLower();
                                    if (textRange.Contains(search)) {
                                        Console.WriteLine($"Found it on slide: #{i + 1} file: {Path.GetFileName(path)}");
                                    }
                                }
                            }
                        }
                    }
                }

                Console.WriteLine("Finished. (hit x to exit)");
                if (Console.ReadKey().Key == ConsoleKey.X)
                    break;
            }
        }

        private static IEnumerable<string> GetAllPowerpoints() {
            var basedir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            var dir = new DirectoryInfo(basedir);
            foreach (var file in dir.GetFiles("*.pptx")) {
                yield return file.FullName;
            }
        }
    }
}
