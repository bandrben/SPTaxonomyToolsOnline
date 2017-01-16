using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;

namespace BandR
{
    public static class ImportFileHelper
    {

        public static bool GetUpdateSimpleDataFromExcelFile(string inputFile, out List<SimpleImportObj> lstObjs, out string msg)
        {
            msg = "";
            lstObjs = new List<SimpleImportObj>();

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(inputFile)))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    var loop = true;
                    int i = 0;
                    while (loop)
                    {
                        i++;

                        // get first cell, being termname
                        var cell1 = worksheet.Cells[i, 1].Value.SafeTrim();

                        if (cell1.IsNull())
                        {
                            loop = false;
                        }
                        else if (i == 1)
                        {
                            // header row, skip
                        }
                        else
                        {
                            var termid = Guid.Parse(worksheet.Cells[i, 2].Value.SafeTrim());
                            var termname = worksheet.Cells[i, 5].Value.SafeTrim();
                            var descr = worksheet.Cells[i, 6].Value.SafeTrim();
                            var isavailfortagging = GenUtil.SafeToBool(worksheet.Cells[i, 7].Value);

                            var labels = new List<string>();
                            int j = 10;
                            while (true)
                            {
                                var cellj = worksheet.Cells[i, j].Value.SafeTrim();

                                if (cellj.IsNull())
                                {
                                    break;
                                }
                                else
                                {
                                    labels.Add(cellj);
                                }

                                j++;
                            }

                            // remove termname from label set, and return distinct only
                            labels.RemoveAll(x => x.Trim().ToLower() == termname.Trim().ToLower());
                            labels = labels.Distinct().ToList();

                            lstObjs.Add(new SimpleImportObj()
                            {
                                descr = descr,
                                isAvailForTagging = isavailfortagging,
                                termId = termid,
                                termName = termname,
                                labels = labels
                            });
                            
                        }

                    } // while
                } // using

            }
            catch (Exception ex)
            {
                msg = ex.ToString();
            }

            return msg == "";
        }

        public static bool GetDataFromExcelFileAdv(string sep, string inputTextFile, out List<TermObjAdv> lstTermObjs, out string msg)
        {
            msg = "";
            lstTermObjs = new List<TermObjAdv>();
            var lines = new List<string>();

            try
            {
                // convert excel data in list of strings(representing lines) of tab separated termparts, since algo for handling that already exists
                using (ExcelPackage package = new ExcelPackage(new FileInfo(inputTextFile)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    var loop = true;
                    int i = 0;
                    while (loop)
                    {
                        i++;

                        var cell1 = worksheet.Cells[i, 1].Value.SafeTrim();

                        if (cell1.IsNull())
                        {
                            loop = false;
                        }
                        else
                        {
                            var termParts = new List<string>();

                            int j = 1;
                            while (true)
                            {
                                var cellj = worksheet.Cells[i, j].Value.SafeTrim();

                                if (cellj.IsNull())
                                {
                                    break;
                                }
                                else
                                {
                                    termParts.Add(cellj);
                                }

                                j++;
                            }

                            lines.Add(string.Join(sep, termParts.ToArray()));
                        }
                    } // while
                } // using

                GetDataAdvWorker(sep, lstTermObjs, lines);

            }
            catch (Exception ex)
            {
                msg = ex.ToString();
            }

            return msg == "";
        }

        public static bool GetDataFromTextFileAdv(string sep, string importTermsPasteBox, string inputTextFile, out List<TermObjAdv> lstTermObjs, out string msg)
        {
            msg = "";
            lstTermObjs = new List<TermObjAdv>();
            var fileText = "";

            try
            {
                if (!importTermsPasteBox.IsNull())
                {
                    fileText = importTermsPasteBox.Trim();
                }
                else
                {
                    fileText = System.IO.File.ReadAllText(inputTextFile);
                }

                fileText = GenUtil.NormalizeEol(fileText);
                var lines = fileText.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Distinct().ToList();

                GetDataAdvWorker(sep, lstTermObjs, lines);

            }
            catch (Exception ex)
            {
                msg = ex.ToString();
            }

            return msg == "";
        }

        private static void GetDataAdvWorker(string sep, List<TermObjAdv> lstTermObjs, List<string> lines)
        {
            foreach (var line in lines)
            {
                var termParts = line.Split(new string[] { sep }, StringSplitOptions.RemoveEmptyEntries).AsEnumerable();

                int i = 0;
                var curPath = "";
                 
                foreach (var termPart in termParts)
                {
                    var curTermPart = termPart;

                    if (curTermPart.IsNull())
                    {
                        continue;
                    }

                    // breakup termpart into [guid]|name|[label0]|[label1]...
                    var termObj = new TermObjAdv();
                    var termId = Guid.NewGuid();
                    var termName = curTermPart.Trim();

                    var termSubParts = curTermPart.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries).AsEnumerable();

                    if (termSubParts.Count() > 1)
                    {
                        int jStart = 0;

                        if (GenUtil.IsGuid(termSubParts.ElementAt(0)))
                        {
                            termId = Guid.Parse(termSubParts.ElementAt(0).Trim());
                            termName = termSubParts.ElementAt(1).Trim();
                            jStart = 2;
                        }
                        else
                        {
                            termName = termSubParts.ElementAt(0).Trim();
                            jStart = 1;
                        }

                        for (int j = jStart; j < termSubParts.Count(); j++)
                        {
                            if (!termSubParts.ElementAt(j).IsNull())
                            {
                                termObj.labels.Add(termSubParts.ElementAt(j).SafeTrim());
                            }
                        }
                    }

                    // fix termname, extract reuse keyword
                    if (termName.ToLower().StartsWith("#reuse") || termName.ToLower().StartsWith("$reuse"))
                    {
                        termObj.isreused = true;
                        termObj.reusebranch = termName.StartsWith("#reuseall") || termName.StartsWith("$reuseall");

                        termName = Regex.Replace(termName, Regex.Escape("#reuseall"), "", RegexOptions.IgnoreCase);
                        termName = Regex.Replace(termName, Regex.Escape("#reuse"), "", RegexOptions.IgnoreCase);
                        termName = Regex.Replace(termName, Regex.Escape("$reuseall"), "", RegexOptions.IgnoreCase);
                        termName = Regex.Replace(termName, Regex.Escape("$reuse"), "", RegexOptions.IgnoreCase);
                    }

                    // trim labels to unique list
                    termObj.labels = termObj.labels.Distinct().ToList();
                    // remove termname from labels
                    termObj.labels.RemoveAll(x => x.ToLower() == termName.ToLower());

                    curPath += termName + ";";

                    termObj.id = termId;
                    termObj.termName = termName;
                    termObj.level = i;
                    termObj.path = curPath.TrimEnd(";".ToCharArray());

                    // add unique term paths only
                    if (!lstTermObjs.Any(x => x.path.ToLower() == termObj.path.ToLower()))
                    {
                        lstTermObjs.Add(termObj);
                    }

                    i++;

                } // foreach
            } // foreach
        }

        public static bool GetDataFromSqlSimple(string dbConnString, string selectStmt, out List<TermObj> lstTermObjs, out string msg)
        {
            msg = "";
            lstTermObjs = new List<TermObj>();

            try
            {
                DataTable dt = null;

                if (!SQLHelper.ExecuteQueryDt(dbConnString, selectStmt, out dt, out msg))
                {
                    throw new Exception(msg);
                }

                if (dt != null && dt.Rows.Count > 0 && dt.Columns.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var termName = dt.Rows[i][0].SafeTrim();
                        var labels = new List<string>();

                        if (!termName.IsNull())
                        {
                            for (int j = 1; j < dt.Columns.Count; j++)
                            {
                                var curLabel = dt.Rows[i][j].SafeTrim();

                                if (!curLabel.IsNull() && !curLabel.IsEqual(termName))
                                {
                                    labels.Add(curLabel);
                                }
                            }

                            lstTermObjs.Add(new TermObj()
                            {
                                termId = Guid.NewGuid(),
                                termName = termName,
                                labels = labels.Distinct().ToList()
                            });
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                msg = ex.ToString();
            }

            return msg == "";
        }

        public static bool GetDataFromExcelFileSimple(string inputFile, out List<TermObj> lstTermObjs, out string msg)
        {
            // term guid not imported
            // terms flat in file, no heirarchy
            // can only import terms flat, and term labels for the imported term
            msg = "";
            lstTermObjs = new List<TermObj>();

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(inputFile)))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    var loop = true;
                    int i = 0;
                    while (loop)
                    {
                        i++;

                        // get first cell, being termname
                        var firstCell = worksheet.Cells[i, 1].Value.SafeTrim();

                        if (firstCell.IsNull())
                        {
                            loop = false;
                        }
                        else
                        {
                            var termName = "";
                            Guid termId;
                            var labels = new List<string>();
                            int j;

                            if (GenUtil.IsGuid(firstCell))
                            {
                                j = 3;
                                termId = GenUtil.SafeToGuid(firstCell).Value;
                                termName = worksheet.Cells[i, 2].Value.SafeTrim();
                            }
                            else
                            {
                                j = 2;
                                termId = Guid.NewGuid();
                                termName = firstCell;
                            }

                            // make sure termname was found
                            if (!termName.IsNull())
                            {
                                // get labels (optional)
                                while (true)
                                {
                                    var cellj = worksheet.Cells[i, j].Value.SafeTrim();

                                    if (cellj.IsNull())
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        labels.Add(cellj);
                                    }

                                    j++;
                                }

                                labels.RemoveAll(x => x.Trim().ToLower() == termName.Trim().ToLower());

                                lstTermObjs.Add(new TermObj()
                                {
                                    termId = termId,
                                    termName = termName,
                                    labels = labels
                                });
                            }
                        }

                    } // while
                } // using

            }
            catch (Exception ex)
            {
                msg = ex.ToString();
            }

            return msg == "";
        }

        public static bool GetDataFromTextFileSimple(string sep, string importTermsPasteBox, string inputTextFile, out List<TermObj> lstTermObjs, out string msg)
        {
            // term guid not imported
            // terms flat in file, no heirarchy
            // can only import terms flat, and term labels for the imported term
            msg = "";
            lstTermObjs = new List<TermObj>();
            var fileText = "";

            try
            {
                if (!importTermsPasteBox.IsNull())
                {
                    fileText = importTermsPasteBox.Trim();
                }
                else
                {
                    fileText = System.IO.File.ReadAllText(inputTextFile);
                }

                fileText = GenUtil.NormalizeEol(fileText);
                var lines = fileText.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Distinct().ToList();

                foreach (var line in lines)
                {
                    // extract termname and optional labels
                    var termId = Guid.NewGuid();
                    var termName = "";
                    var labels = new List<string>();

                    if (line.Contains(sep))
                    {
                        var parts = line.Split(new string[] { sep }, StringSplitOptions.RemoveEmptyEntries).Distinct();

                        if (GenUtil.IsGuid(parts.ElementAt(0)))
                        {
                            // guid is first item, termname must be second (thus 2 item minimum)
                            if (parts.Count() >= 2)
                            {
                                termId = GenUtil.SafeToGuid(parts.ElementAt(0)).Value;
                                termName = parts.ElementAt(1).Trim();
                                labels = parts.Skip(2).Where(x => x.Trim().Length > 0).Select(x => x.Trim()).Distinct().ToList();
                            }
                            else
                            {
                                // termname not found, don't import this line
                                termName = "";
                            }
                        }
                        else
                        {
                            termName = parts.ElementAt(0).Trim();
                            labels = parts.Skip(1).Where(x => x.Trim().Length > 0).Select(x => x.Trim()).Distinct().ToList();
                        }

                        labels.RemoveAll(x => x.Trim().ToLower() == termName.Trim().ToLower());

                    }
                    else
                    {
                        termName = line.Trim();
                    }

                    if (!termName.IsNull())
                    {
                        lstTermObjs.Add(new TermObj()
                        {
                            termId = termId,
                            termName = termName,
                            labels = labels
                        });
                    }
                }

            }
            catch (Exception ex)
            {
                msg = ex.ToString();
            }

            return msg == "";
        }

    }
}
