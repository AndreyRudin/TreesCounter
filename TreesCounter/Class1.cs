using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System.Linq;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Windows;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;

namespace TreesCounter
{
    public class Command : IExtensionApplication
    {

        [CommandMethod("CountTrees")]
        public void Helloworld()
        {
            //вынести допуски по вертикали и горизонтали в интерфейс
            //ошибочный текст сгруппировать и отсортировать по алфавиту хотя бы

            List<DBText> allTexts = new List<DBText>();
            List<DBText> secondTexts = new List<DBText>();
            List<DBText> othersTexts = new List<DBText>();
            List<TreeData> firstTrees = new List<TreeData>();
            List<TreeData> firstAndSecondTrees = new List<TreeData>();
            List<TreeData> firstAndSecondTreeGeneral = new List<TreeData>();
            List<TreeData> correctFirstAndSecondTrees = new List<TreeData>();
            List<TreeData> errorFirstAndSecondTrees = new List<TreeData>();

            Document acDoc = acadApp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;

            PromptSelectionOptions options = new PromptSelectionOptions();
            options.SingleOnly = false;

            PromptSelectionResult selRes = ed.GetSelection(options);

            if (selRes.Status == PromptStatus.OK)
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject item in selRes.Value)
                    {
                        Entity acEnt = acTrans.GetObject(item.ObjectId, OpenMode.ForRead) as Entity;
                        var text = acEnt as DBText;
                        if (text != null)
                        {
                            allTexts.Add(text);
                        }
                    }

                    foreach (DBText text in allTexts)
                    {
                        var str = text.TextString;

                        if (str == "сух.")
                        {
                            firstTrees.Add(new TreeData(text));
                        }

                        Regex regexTextNum = new Regex(@"[А-ЯЁа-яё]+-\d+");
                        MatchCollection matchesTextNum = regexTextNum.Matches(str);
                        if (matchesTextNum.Count > 0 && matchesTextNum[0].Value == str)
                        {
                            var findTree = firstTrees.Find(tree => tree.FirstName == str);
                            if (findTree != null)
                            {
                                findTree.Count += 1;
                            }
                            else
                            {
                                firstTrees.Add(new TreeData(text));
                            }
                            continue;
                        }

                        Regex regexNumText = new Regex(@"\d+-[А-ЯЁа-яё]+");
                        MatchCollection matchesNumText = regexNumText.Matches(str);
                        if (matchesNumText.Count > 0 && matchesNumText[0].Value == str)
                        {
                            firstAndSecondTrees.Add(new TreeData(text));
                            continue;
                        }

                        Regex regexNumNum = new Regex(@"\d+-\d+");
                        MatchCollection matchesNumNum = regexNumNum.Matches(str);
                        if (matchesNumNum.Count > 0 && matchesNumNum[0].Value == str)
                        {
                            secondTexts.Add(text);
                            continue;
                        }

                        Regex regexNum = new Regex(@"\d+");
                        MatchCollection matchesNum = regexNum.Matches(str);
                        if (matchesNum.Count > 0 && matchesNum[0].Value == str)
                        {
                            secondTexts.Add(text);
                            continue;
                        }

                        othersTexts.Add(text);
                    }
                    foreach (var tree in firstAndSecondTrees)
                    {
                        var firstTextX = tree.FirstText.Position.X;
                        var firstTextY = tree.FirstText.Position.Y;
                        foreach (var dbtext in secondTexts)
                        {
                            var secondTextX = dbtext.Position.X;
                            var secondTextY = dbtext.Position.Y;
                            if ((Math.Abs(firstTextX - secondTextX) <= 1.17 && Math.Abs(firstTextY - secondTextY) <= 1.5) &&
                                (firstTextY > secondTextY))
                            {
                                tree.SecondTexts.Add(dbtext);
                            }
                        }
                    }

                    errorFirstAndSecondTrees = firstAndSecondTrees.Where(tree => tree.CountSecondText != 1).ToList();
                    correctFirstAndSecondTrees = firstAndSecondTrees.Where(tree => tree.CountSecondText == 1).ToList();

                    foreach (var correctFirstAndSecondTree in correctFirstAndSecondTrees)
                    {
                        var findTree = firstAndSecondTreeGeneral
                            .Find(tree => tree.FirstName == correctFirstAndSecondTree.FirstName && tree.SecondName == correctFirstAndSecondTree.SecondName);
                        if (findTree != null)
                        {
                            findTree.Count += 1;
                        }
                        else
                        {
                            firstAndSecondTreeGeneral.Add(correctFirstAndSecondTree);
                        }
                    }
                    var listReports = new List<string>();
                    listReports.Add($"Корректные деревья с одной строкой: всего типов {firstTrees.Count}, всего штук: {firstTrees.Select(t => t.Count).Sum()}");
                    listReports.AddRange(firstTrees.OrderBy(t => t.FirstName).Select(t => $"имя: {t.FirstName} кол-во: {t.Count}").ToList());
                    listReports.Add("");

                    listReports.Add($"Корректные деревья с двумя строками : всего типов {correctFirstAndSecondTrees.Count}, всего штук: {correctFirstAndSecondTrees.Select(t => t.Count).Sum()}");
                    listReports.AddRange(correctFirstAndSecondTrees.OrderBy(t => t.FirstName).ThenBy(t => t.SecondName).Select(t => $"верхнее имя: {t.FirstName} нижнее имя: {t.SecondName} кол-во: {t.Count}").ToList());
                    listReports.Add("");

                    listReports.Add("Некорректные деревья с двумя строками");
                    foreach (var errorTree in errorFirstAndSecondTrees)
                    {
                        listReports.Add($"верхнее имя: {errorTree.FirstName} положение верхнего имени {errorTree.FirstTextPoint} количество нижнего текста {errorTree.CountSecondText}");
                        foreach (var secondText in errorTree.SecondTexts)
                        {
                            listReports.Add($"\tнижнее имя: {secondText.TextString} положение нижнего имени {secondText.Position}");
                        }
                    }
                    listReports.Add("");

                    listReports.Add("Текст не прошеднший фильтрацию");
                    listReports.AddRange(othersTexts.OrderBy(t => t.TextString).Select(t => $"Текст: {t.TextString} место на чертеже: {t.Position}").ToList());
                    listReports.Add("");

                    string allReport = string.Join("\n", listReports);

                    string now = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                    string docDir = @"C:\ProgramData";
                    string nameFile = string.Concat("Report", now, ".txt");
                    string PathToReport = Path.Combine(docDir, nameFile);

                    File.WriteAllText(PathToReport, allReport);
                    Process.Start(PathToReport);

                    acTrans.Commit();
                }
            }
        }

        public void Initialize()
        {
        }

        public void Terminate()
        {
        }
    }
}
