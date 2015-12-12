using BandR;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace SPTaxonomyToolsOnline
{
    public partial class Form1 : System.Windows.Forms.Form
    {

        private AboutForm aboutForm = null;
        private BackgroundWorker bgw = null;
        private int statusWindowOutputBatchSize = GenUtil.SafeToInt(ConfigurationManager.AppSettings["statusWindowOutputBatchSize"]);
        private bool showFullErrMsgs = GenUtil.SafeToBool(ConfigurationManager.AppSettings["showFullErrMsgs"]);



        public Form1()
        {
            InitializeComponent();

            toolStripStatusLabel1.Text = "";

            this.FormClosed += Form1_FormClosed;

            LoadSettingsFromRegistry();

            imageBandR.Visible = true;
            imageBandRwait.Visible = false;

            rbExportTypeSimple.Checked = true;
            rbExportFormatText.Checked = true;

            rbImportTypeSimple.Checked = true;
            rbImportSourceText.Checked = true;
        }




        private ICredentials BuildCreds()
        {
            var userName = tbUsername.Text.Trim();
            var passWord = tbPassword.Text.Trim();
            var domain = tbDomain.Text.Trim();

            if (!cbIsSPOnline.Checked)
            {
                return new NetworkCredential(userName, passWord, domain);
            }
            else
            {
                return new SharePointOnlineCredentials(userName, GenUtil.BuildSecureString(passWord));
            }
        }

        private void ctx_ExecutingWebRequest_FixForMixedMode(object sender, WebRequestEventArgs e)
        {
            // to support mixed mode auth
            e.WebRequestExecutor.RequestHeaders.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
        }

        private void FixCtxForMixedMode(ClientContext ctx)
        {
            if (!cbIsSPOnline.Checked)
            {
                ctx.ExecutingWebRequest += new EventHandler<WebRequestEventArgs>(ctx_ExecutingWebRequest_FixForMixedMode);
            }
        }




        private void LoadSettingsFromRegistry()
        {
            var msg = "";
            var json = "";

            if (RegistryHelper.GetRegStuff(out json, out msg) && !json.IsNull())
            {
                var obj = JsonExtensionMethod.FromJson<CustomRegistrySettings>(json);

                tbSiteUrl.Text = obj.siteUrl;
                tbUsername.Text = obj.userName;
                tbPassword.Text = obj.passWord;
                tbDomain.Text = obj.domain;
                cbIsSPOnline.Checked = GenUtil.SafeToBool(obj.isSPOnline);

                tbTermGroup.Text = obj.termGroup;
                tbTermSet.Text = obj.termSet;
                tbTermStore.Text = obj.termStore;

                tbTermGroupID.Text = obj.termGroupID;
                tbTermSetID.Text = obj.termSetID;
                tbTermStoreID.Text = obj.termStoreID;

                tbExportFilePath.Text = obj.exportFilePath;

                tbImportSourceFilePath.Text = obj.importSourceFilePath;
                tbImportSeparator.Text = GenUtil.NVL(obj.importSeparator, "\\t");
                tbImportDbConnString.Text = obj.importDbConnString;
                tbImportSelectStmt.Text = obj.importSelectStmt;

                tbUpdateTermsSourceFilePath.Text = obj.updateTermsSourceFilePath;

            }
        }

        private void SaveSettingsToRegistry()
        {
            var msg = "";

            var obj = new CustomRegistrySettings
            {
                siteUrl = tbSiteUrl.Text.Trim(),
                userName = tbUsername.Text.Trim(),
                passWord = tbPassword.Text.Trim(),
                domain = tbDomain.Text.Trim(),
                isSPOnline = cbIsSPOnline.Checked ? "1" : "0",

                termGroup = tbTermGroup.Text.Trim(),
                termSet = tbTermSet.Text.Trim(),
                termStore = tbTermStore.Text.Trim(),

                termGroupID = tbTermGroupID.Text.Trim(),
                termSetID = tbTermSetID.Text.Trim(),
                termStoreID = tbTermStoreID.Text.Trim(),

                exportFilePath = tbExportFilePath.Text.Trim(),

                importSourceFilePath = tbImportSourceFilePath.Text.Trim(),
                importSeparator = tbImportSeparator.Text.Trim(),
                importDbConnString = tbImportDbConnString.Text.Trim(),
                importSelectStmt = tbImportSelectStmt.Text.Trim(),

                updateTermsSourceFilePath = tbUpdateTermsSourceFilePath.Text.Trim()
            };

            var json = JsonExtensionMethod.ToJson(obj);

            RegistryHelper.SaveRegStuff(json, out msg);
        }





        void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (aboutForm != null)
            {
                aboutForm.Dispose();
            }

            SaveSettingsToRegistry();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (aboutForm == null)
            {
                aboutForm = new AboutForm();
            }

            aboutForm.Show();
            aboutForm.Focus();
        }

        private void DisableFormControls()
        {
            toolStripStatusLabel1.Text = "Running...";

            imageBandR.Visible = false;
            imageBandRwait.Visible = true;

            btnStartClearTermSet.Enabled = false;
            btnStartExport.Enabled = false;
            btnStartImport.Enabled = false;
            btnStartTestConnection.Enabled = false;
            btnLoadMMD.Enabled = false;

            lnkClear.Enabled = false;
            lnkExport.Enabled = false;
        }

        private void EnableFormControls()
        {
            toolStripStatusLabel1.Text = "";

            imageBandR.Visible = true;
            imageBandRwait.Visible = false;

            btnStartClearTermSet.Enabled = true;
            btnStartExport.Enabled = true;
            btnStartImport.Enabled = true;
            btnStartTestConnection.Enabled = true;
            btnLoadMMD.Enabled = true;

            lnkClear.Enabled = true;
            lnkExport.Enabled = true;
        }





        private void btnStartTestConnection_Click(object sender, EventArgs e)
        {
            DisableFormControls();
            InitCoutBuffer();
            tbStatus.Text = "";

            bgw = new BackgroundWorker();
            bgw.DoWork += new DoWorkEventHandler(bgw_TestConnection);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_TestConnection_End);
            bgw.ProgressChanged += new ProgressChangedEventHandler(BgwReportProgress);
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync();
        }

        private void bgw_TestConnection(object sender, DoWorkEventArgs e)
        {
            try
            {
                tcout("Testing SharePoint Connection:");

                tcout("SiteUrl", tbSiteUrl.Text.Trim());

                var targetSite = new Uri(tbSiteUrl.Text.Trim());

                using (ClientContext ctx = new ClientContext(targetSite))
                {
                    ctx.Credentials = BuildCreds();
                    FixCtxForMixedMode(ctx);

                    Web web = ctx.Web;
                    ctx.Load(web, w => w.Title);
                    ctx.ExecuteQuery();
                    tcout("Site loaded OK", web.Title);
                }

            }
            catch (Exception ex)
            {
                tcout("ERROR", GetExcMsg(ex));
            }
        }

        private void bgw_TestConnection_End(object sender, RunWorkerCompletedEventArgs e)
        {
            FlushCoutBuffer();
            EnableFormControls();
        }






        private void LoadTStoreTGroupTSet(ClientContext ctx, TaxonomySession session, ref TermStore store, ref TermGroup group, ref TermSet set)
        {
            // load termstore
            if (!GenUtil.IsNull(tbTermStore.Text.Trim()))
            {
                store = session.TermStores.GetByName(tbTermStore.Text.Trim());
            }
            else if (!GenUtil.IsNull(tbTermStoreID.Text.Trim()))
            {
                store = session.TermStores.GetById(GenUtil.SafeToGuid(tbTermStoreID.Text.Trim()).Value);
            }
            else
            {
                throw new Exception("Term Store name or ID missing.");
            }

            ctx.Load(store);
            ctx.ExecuteQuery();

            tcout("Termstore loaded", store.Name);

            // load termgroup
            if (!GenUtil.IsNull(tbTermGroup.Text.Trim()))
            {
                group = store.Groups.GetByName(tbTermGroup.Text.Trim());
            }
            else if (!GenUtil.IsNull(tbTermGroupID.Text.Trim()))
            {
                group = store.Groups.GetById(GenUtil.SafeToGuid(tbTermGroupID.Text.Trim()).Value);
            }
            else
            {
                throw new Exception("Term Group name or ID missing.");
            }

            ctx.Load(group);
            ctx.ExecuteQuery();

            tcout("Termgroup loaded", group.Name);

            // load termset
            if (!GenUtil.IsNull(tbTermSet.Text.Trim()))
            {
                set = group.TermSets.GetByName(tbTermSet.Text.Trim());
            }
            else if (!GenUtil.IsNull(tbTermSetID.Text.Trim()))
            {
                set = group.TermSets.GetById(GenUtil.SafeToGuid(tbTermSetID.Text.Trim()).Value);
            }
            else
            {
                throw new Exception("Term Set name or ID missing.");
            }

            ctx.Load(set);
            ctx.ExecuteQuery();

            tcout("Termset loaded", set.Name);
        }







        private void btnStartExport_Click(object sender, EventArgs e)
        {
            DisableFormControls();
            InitCoutBuffer();
            tbStatus.Text = "";

            bgw = new BackgroundWorker();
            bgw.DoWork += new DoWorkEventHandler(bgw_btnStartExport);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_btnStartExport_End);
            bgw.ProgressChanged += new ProgressChangedEventHandler(BgwReportProgress);
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync();
        }

        private void bgw_btnStartExport(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (rbExportTypeSimple.Checked)
                {
                    GetAllTermsSimple();
                }
                else
                {
                    GetAllTermsAdv();
                }

            }
            catch (Exception ex)
            {
                tcout("ERROR", GetExcMsg(ex));
            }
        }

        private void GetAllTermsAdv()
        {
            tcout("SiteUrl", tbSiteUrl.Text);
            tcout("UserName", tbUsername.Text);
            tcout("Domain", tbDomain.Text);
            tcout("IsSPOnline", cbIsSPOnline.Checked);
            tcout("TermStore", tbTermStore.Text);
            tcout("TermGroup", tbTermGroup.Text);
            tcout("TermSet", tbTermSet.Text);
            tcout("TermStoreID", tbTermStoreID.Text);
            tcout("TermGroupID", tbTermGroupID.Text);
            tcout("TermSetID", tbTermSetID.Text);
            tcout("ExportType", rbExportTypeSimple.Checked ? "Simple" : "Advanced");
            tcout("ExportFormat", rbExportFormatText.Checked ? "Text" : (rbExportFormatExcel.Checked ? "Excel" : "Screen"));
            tcout("ExportTermGuids", cbExportTermIds.Checked);
            tcout("ExportTermLabels", cbExportTermLabels.Checked);
            tcout("---------------------------------------------");

            var subSeparator = "|";

            var targetSite = new Uri(tbSiteUrl.Text.Trim());

            using (ClientContext ctx = new ClientContext(targetSite))
            {
                ctx.Credentials = BuildCreds();
                FixCtxForMixedMode(ctx);

                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQuery();
                tcout("Site loaded", web.Title);

                var lcid = System.Globalization.CultureInfo.CurrentCulture.LCID;
                var session = TaxonomySession.GetTaxonomySession(ctx);

                TermStore store = null;
                TermGroup group = null;
                TermSet set = null;

                LoadTStoreTGroupTSet(ctx, session, ref store, ref group, ref set);

                // get all terms
                var allMmdTerms = set.GetAllTerms();

                ctx.Load(allMmdTerms, a =>
                    a.Include(
                        b => b.Id,
                        b => b.Name,
                        b => b.PathOfTerm,
                        b => b.Labels.Include(
                            c => c.IsDefaultForLanguage,
                            c => c.Value)));
                ctx.ExecuteQuery();

                tcout("All terms loaded", allMmdTerms.Count);

                var lstTerms = new List<TermObjAdv>();

                foreach (var curMmdTerm in allMmdTerms)
                {
                    var termObjAdv = new TermObjAdv();
                    termObjAdv.id = curMmdTerm.Id;
                    termObjAdv.termName = GenUtil.MmdDenormalize(curMmdTerm.Name);
                    termObjAdv.path = GenUtil.MmdDenormalize(curMmdTerm.PathOfTerm);

                    var labels = new List<string>();
                    foreach (var label in curMmdTerm.Labels)
                    {
                        if (!label.IsDefaultForLanguage)
                        {
                            labels.Add(GenUtil.MmdDenormalize(label.Value));
                        }
                    }

                    termObjAdv.labels = labels;
                    termObjAdv.level = termObjAdv.path.ToCharArray().Count(x => x == ';');

                    lstTerms.Add(termObjAdv);
                }

                var sb = new StringBuilder();

                foreach (var term in lstTerms.OrderBy(x => x.level))
                {
                    var pathParts = term.path.Split(";".ToCharArray());

                    var line = "";
                    for (int i = 0; i <= term.level; i++)
                    {
                        var tmp = lstTerms.FirstOrDefault(x => x.level == i && x.termName.Trim().ToLower() == pathParts[i].Trim().ToLower());

                        if (cbExportTermIds.Checked)
                        {
                            line += tmp.id + subSeparator;
                        }

                        line += tmp.termName;

                        if (cbExportTermLabels.Checked && tmp.labels.Any())
                        {
                            line += subSeparator + string.Join(subSeparator, tmp.labels.ToArray());
                        }

                        line += "\t";
                    }
                    sb.AppendLine(line.Trim());
                }

                if (sb.Length == 0)
                {
                    tcout("No terms found to export.");
                }
                else
                {
                    if (rbExportFormatExcel.Checked)
                    {
                        // export to Excel file
                        var path = AppDomain.CurrentDomain.BaseDirectory;
                        if (!tbExportFilePath.Text.IsNull())
                            path = tbExportFilePath.Text.Trim();
                        path = path.CombineFS(GenUtil.CleanFilenameForFS("allterms_" + set.Name + "_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")) + ".xlsx");

                        using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
                        {
                            var sheet = package.Workbook.Worksheets.Add("Sheet1");
                            sheet.DefaultColWidth = 20;

                            int i = 1;

                            foreach (var line in sb.ToString().Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                            {
                                var curLineObjs = line.Split("\t".ToCharArray());

                                for (int j = 1; j <= curLineObjs.Length; j++)
                                {
                                    sheet.Cells[i, j].Value = curLineObjs[j - 1];
                                }

                                i++;
                            }

                            package.Save();
                            tcout("All terms exported", path);
                        }
                    }
                    else if (rbExportFormatText.Checked)
                    {
                        // export terms to tab delim file
                        var path = AppDomain.CurrentDomain.BaseDirectory;
                        if (!tbExportFilePath.Text.IsNull())
                            path = tbExportFilePath.Text.Trim();
                        path = path.CombineFS(GenUtil.CleanFilenameForFS("allterms_" + set.Name + "_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")) + ".txt");

                        System.IO.File.WriteAllText(path, sb.ToString());
                        tcout("All terms exported", path);
                    }
                    else
                    {
                        // export to screen
                        tcout("---------------------------------------------");
                        foreach (var line in sb.ToString().Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                        {
                            tcout(line);
                        }
                        tcout("---------------------------------------------");
                    }
                }
            }
        }

        private void GetAllTermsSimple()
        {
            tcout("SiteUrl", tbSiteUrl.Text);
            tcout("UserName", tbUsername.Text);
            tcout("Domain", tbDomain.Text);
            tcout("IsSPOnline", cbIsSPOnline.Checked);
            tcout("TermStore", tbTermStore.Text);
            tcout("TermGroup", tbTermGroup.Text);
            tcout("TermSet", tbTermSet.Text);
            tcout("TermStoreID", tbTermStoreID.Text);
            tcout("TermGroupID", tbTermGroupID.Text);
            tcout("TermSetID", tbTermSetID.Text);
            tcout("ExportFormat", rbExportFormatText.Checked ? "Text" : (rbExportFormatExcel.Checked ? "Excel" : "Screen"));
            tcout("---------------------------------------------");

            var targetSite = new Uri(tbSiteUrl.Text.Trim());

            using (ClientContext ctx = new ClientContext(targetSite))
            {
                ctx.Credentials = BuildCreds();
                FixCtxForMixedMode(ctx);

                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQuery();
                tcout("Site loaded", web.Title);

                var lcid = System.Globalization.CultureInfo.CurrentCulture.LCID;
                var session = TaxonomySession.GetTaxonomySession(ctx);

                TermStore store = null;
                TermGroup group = null;
                TermSet set = null;

                LoadTStoreTGroupTSet(ctx, session, ref store, ref group, ref set);

                // get all terms
                var allMmdTerms = set.GetAllTerms();

                ctx.Load(allMmdTerms, a =>
                    a.Include(
                        b => b.Id,
                        b => b.Name,
                        b => b.PathOfTerm,
                        b => b.Description,
                        b => b.IsAvailableForTagging,
                        b => b.IsReused,
                        b => b.IsSourceTerm,
                        b => b.Labels.Include(
                            c => c.IsDefaultForLanguage,
                            c => c.Value),
                        b => b.Parent.Id));
                ctx.ExecuteQuery();

                tcout("All terms loaded", allMmdTerms.Count);

                var simpleExportObjs = new List<SimpleExportObj>();
                int i = 0;

                foreach (var curMmdTerm in allMmdTerms)
                {
                    i++;

                    var parentId = "";
                    var parent = curMmdTerm.Parent;
                    if (parent != null && parent.ServerObjectIsNull.HasValue && !parent.ServerObjectIsNull.Value)
                    {
                        parentId = parent.Id.ToString();
                    }

                    var labels = new List<string>();
                    foreach (var label in curMmdTerm.Labels)
                    {
                        if (!label.IsDefaultForLanguage)
                        {
                            labels.Add(GenUtil.MmdDenormalize(label.Value));
                        }
                    }

                    simpleExportObjs.Add(new SimpleExportObj
                    {
                        i = i.ToString(),
                        id = curMmdTerm.Id.ToString(),
                        parentId = parentId,
                        pathOfTerm = GenUtil.MmdDenormalize(curMmdTerm.PathOfTerm),
                        name = GenUtil.MmdDenormalize(curMmdTerm.Name),
                        description = curMmdTerm.Description.Replace("\t", "\\t"),
                        isAvailableForTagging = curMmdTerm.IsAvailableForTagging.ToString(),
                        isReused = curMmdTerm.IsReused.ToString(),
                        isSourceTerm = curMmdTerm.IsSourceTerm.ToString(),
                        labels = string.Join("\t", labels.ToArray()).Trim()
                    });
                }

                var sb = new StringBuilder();
                foreach (var curSimpleExportObj in simpleExportObjs.OrderBy(x => x.level))
                {
                    sb.AppendLine(
                        string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}",
                            curSimpleExportObj.i,
                            curSimpleExportObj.id,
                            curSimpleExportObj.parentId,
                            curSimpleExportObj.pathOfTerm,
                            curSimpleExportObj.name,
                            curSimpleExportObj.description,
                            curSimpleExportObj.isAvailableForTagging,
                            curSimpleExportObj.isReused,
                            curSimpleExportObj.isSourceTerm,
                            curSimpleExportObj.labels));
                }

                if (sb.Length == 0)
                {
                    tcout("No terms found to export.");
                }
                else
                {
                    if (rbExportFormatExcel.Checked)
                    {
                        // export to Excel file
                        var path = AppDomain.CurrentDomain.BaseDirectory;
                        if (!tbExportFilePath.Text.IsNull())
                            path = tbExportFilePath.Text.Trim();
                        path = path.CombineFS(GenUtil.CleanFilenameForFS("allterms_" + set.Name + "_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")) + ".xlsx");
                        
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
                        {
                            var sheet = package.Workbook.Worksheets.Add("Sheet1");
                            sheet.DefaultColWidth = 20;

                            var k = 1;
                            sheet.Cells[1, k++].Value = "Counter";
                            sheet.Cells[1, k++].Value = "TermId";
                            sheet.Cells[1, k++].Value = "ParentTermId";
                            sheet.Cells[1, k++].Value = "TermPath";
                            sheet.Cells[1, k++].Value = "TermName";
                            sheet.Cells[1, k++].Value = "Descr";
                            sheet.Cells[1, k++].Value = "IsAvailForTagging";
                            sheet.Cells[1, k++].Value = "IsReused";
                            sheet.Cells[1, k++].Value = "IsSourceTerm";
                            sheet.Cells[1, k++].Value = "TermLabels";

                            sheet.Cells[1, 1, 1, k].Style.Font.Bold = true;

                            i = 2;

                            foreach (var line in sb.ToString().Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                            {
                                var curLineObjs = line.Split("\t".ToCharArray());

                                for (int j = 1; j <= curLineObjs.Length; j++)
                                {
                                    sheet.Cells[i, j].Value = curLineObjs[j - 1];
                                }

                                i++;
                            }

                            package.Save();
                            tcout("All terms exported", path);
                        }
                    }
                    else if (rbExportFormatText.Checked)
                    {
                        // export terms to tab delim file
                        var path = AppDomain.CurrentDomain.BaseDirectory;
                        if (!tbExportFilePath.Text.IsNull())
                            path = tbExportFilePath.Text.Trim();
                        path = path.CombineFS(GenUtil.CleanFilenameForFS("allterms_" + set.Name + "_" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")) + ".txt");

                        var header = "Counter\tTermId\tParentTermId\tTermPath\tTermName\tDescr\tIsAvailForTagging\tIsReused\tIsSourceTerm\tTermLabels\r\n";
                        System.IO.File.WriteAllText(path, header + sb.ToString());
                        tcout("All terms exported", path);
                    }
                    else
                    {
                        // export to screen
                        tcout("---------------------------------------------");
                        tcout("Counter\tTermId\tParentTermId\tTermPath\tTermName\tDescr\tIsAvailForTagging\tIsReused\tIsSourceTerm\tTermLabels");

                        foreach (var line in sb.ToString().Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                        {
                            tcout(line);
                        }

                        tcout("---------------------------------------------");
                    }
                }
            }
        }

        private void bgw_btnStartExport_End(object sender, RunWorkerCompletedEventArgs e)
        {
            FlushCoutBuffer();
            SaveLogToFile("ExportTerms");
            EnableFormControls();
        }












        private void btnStartImport_Click(object sender, EventArgs e)
        {
            DialogResult dgResult = MessageBox.Show("Are you sure?", "Import into Term Set", MessageBoxButtons.YesNo);

            if (dgResult != DialogResult.Yes)
            {
                cout("Canceled");
                return;
            }

            DisableFormControls();
            InitCoutBuffer();
            tbStatus.Text = "";

            bgw = new BackgroundWorker();
            bgw.DoWork += new DoWorkEventHandler(bgw_btnStartImport);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_btnStartImport_End);
            bgw.ProgressChanged += new ProgressChangedEventHandler(BgwReportProgress);
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync();
        }

        private void bgw_btnStartImport(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (rbImportTypeSimple.Checked)
                {
                    ImportTermsSimple();
                }
                else
                {
                    ImportTermsAdv();
                }

            }
            catch (Exception ex)
            {
                tcout("ERROR", GetExcMsg(ex));
            }
        }

        private void ImportTermsAdv()
        {
            var msg = "";

            tcout("SiteUrl", tbSiteUrl.Text);
            tcout("UserName", tbUsername.Text);
            tcout("Domain", tbDomain.Text);
            tcout("IsSPOnline", cbIsSPOnline.Checked);
            tcout("TermStore", tbTermStore.Text);
            tcout("TermGroup", tbTermGroup.Text);
            tcout("TermSet", tbTermSet.Text);
            tcout("TermStoreID", tbTermStoreID.Text);
            tcout("TermGroupID", tbTermGroupID.Text);
            tcout("TermSetID", tbTermSetID.Text);

            tcout("ImportType", rbImportTypeSimple.Checked ? "Simple" : "Advanced");
            tcout("SourceType", rbImportSourceText.Checked ? "Text" : (rbImportSourceExcel.Checked ? "Excel" : "SQL"));
            tcout("AppendNewLabelsToExistingTerms", cbAppendNewLabelsToExistingTerms.Checked);
            tcout("SourceFilePath", tbImportSourceFilePath.Text);
            tcout("TermSeparator", tbImportSeparator.Text);
            tcout("---------------------------------------------");

            var tmpSep = tbImportSeparator.Text.Trim() == "\\t" ? "\t" : tbImportSeparator.Text.Trim();

            // get terms and labels from source
            var lstImportTermObjs = new List<TermObjAdv>();

            if (rbImportSourceText.Checked)
            {
                if (!ImportFileHelper.GetDataFromTextFileAdv(tmpSep, tbImportSourceFilePath.Text.Trim(), out lstImportTermObjs, out msg))
                {
                    tcout("ERROR extracting data from text file", msg);
                    return;
                }
            }
            else if (rbImportSourceExcel.Checked)
            {
                if (!ImportFileHelper.GetDataFromExcelFileAdv(tmpSep, tbImportSourceFilePath.Text.Trim(), out lstImportTermObjs, out msg))
                {
                    tcout("ERROR extracting data from excel file", msg);
                    return;
                }
            }

            if (!lstImportTermObjs.Any())
            {
                tcout("Found no terms to import");
                return;
            }

            // begin import terms
            var targetSite = new Uri(tbSiteUrl.Text.Trim());

            using (ClientContext ctx = new ClientContext(targetSite))
            {
                ctx.Credentials = BuildCreds();
                FixCtxForMixedMode(ctx);

                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQuery();
                tcout("Site loaded", web.Title);

                var lcid = System.Globalization.CultureInfo.CurrentCulture.LCID;
                var session = TaxonomySession.GetTaxonomySession(ctx);

                TermStore store = null;
                TermGroup group = null;
                TermSet set = null;

                LoadTStoreTGroupTSet(ctx, session, ref store, ref group, ref set);

                // get all terms in termset, will be checking paths
                var mmdTerms = set.GetAllTerms();

                if (cbAppendNewLabelsToExistingTerms.Checked)
                {
                    ctx.Load(mmdTerms, a =>
                        a.Include(
                            b => b.Id,
                            b => b.Name,
                            b => b.PathOfTerm,
                            b => b.Labels));
                }
                else
                {
                    ctx.Load(mmdTerms, a =>
                        a.Include(
                            b => b.Id,
                            b => b.Name,
                            b => b.PathOfTerm));
                }

                ctx.ExecuteQuery();

                // convert to list of terms so new terms can be added to collection
                var lstMmdTerms = mmdTerms.ToList<Term>();

                // start at level0, then proceed deeper
                for (int level = 0; level <= lstImportTermObjs.Max(x => x.level); level++)
                {
                    // get cur level terms
                    var curLevelTermObjs = lstImportTermObjs.Where(x => x.level == level);

                    foreach (var levelTermObj in curLevelTermObjs)
                    {
                        var curLevelTermObj = levelTermObj;

                        tcout("checking term name", curLevelTermObj.termName, "path", curLevelTermObj.path);

                        if (curLevelTermObj.isreused)
                        {
                            // add reused term
                            var sourceTerm = store.GetTerm(curLevelTermObj.id);

                            if (level == 0)
                            {
                                // add term to termset
                                try
                                {
                                    var newTerm = set.ReuseTerm(sourceTerm, curLevelTermObj.reusebranch);
                                    ctx.ExecuteQuery();

                                    ctx.Load(newTerm, a => a.Name, a => a.PathOfTerm);
                                    ctx.ExecuteQuery();
                                    lstMmdTerms.Add(newTerm);

                                    tcout(" -- ", "term reused", newTerm.PathOfTerm);

                                }
                                catch (Exception ex)
                                {
                                    tcout(" -- ", "ERROR adding reused term", curLevelTermObj.id, GetExcMsg(ex));
                                }
                            }
                            else
                            {
                                // find parent term, add term to term
                                var parentTermPath = curLevelTermObj.path.Substring(0, curLevelTermObj.path.LastIndexOf(';'));
                                var parentTerm = lstMmdTerms.FirstOrDefault(x => GenUtil.MmdDenormalize(x.PathOfTerm).ToLower() == GenUtil.MmdDenormalize(parentTermPath).ToLower());

                                if (parentTerm == null)
                                {
                                    tcout(" -- ", "parent term not found, cannot reuse term");
                                }
                                else
                                {
                                    try
                                    {
                                        var newTerm = parentTerm.ReuseTerm(sourceTerm, curLevelTermObj.reusebranch);
                                        ctx.ExecuteQuery();

                                        ctx.Load(newTerm, a => a.Name, a => a.PathOfTerm);
                                        ctx.ExecuteQuery();
                                        lstMmdTerms.Add(newTerm);

                                        tcout(" -- ", "term reused", newTerm.PathOfTerm);

                                    }
                                    catch (Exception ex)
                                    {
                                        tcout(" -- ", "ERROR adding reused term", curLevelTermObj.id, GetExcMsg(ex));
                                    }
                                }
                            }

                        }
                        else
                        {
                            // check if curterm exists in termset, comparing paths
                            var termMatch = lstMmdTerms.FirstOrDefault(x => GenUtil.MmdDenormalize(x.PathOfTerm).ToLower() == GenUtil.MmdDenormalize(curLevelTermObj.path).ToLower());

                            if (termMatch != null)
                            {
                                // term found, optionally add new labels
                                if (cbAppendNewLabelsToExistingTerms.Checked)
                                {
                                    if (curLevelTermObj.labels.Any())
                                    {
                                        foreach (var label in curLevelTermObj.labels)
                                        {
                                            var curLabel = label;

                                            var found = false;
                                            foreach (var matchLabel in termMatch.Labels)
                                            {
                                                if (matchLabel.Value.IsEqual(curLabel))
                                                {
                                                    found = true;
                                                    break;
                                                }
                                            }

                                            if (!found)
                                            {
                                                // create label
                                                try
                                                {
                                                    termMatch.CreateLabel(curLabel, lcid, false);
                                                    ctx.ExecuteQuery();
                                                    tcout(" -- ", "label created", curLabel);
                                                }
                                                catch (Exception ex)
                                                {
                                                    tcout(" -- ", "ERROR creating label", curLabel, GetExcMsg(ex));
                                                }
                                            }
                                        }
                                    }
                                }

                            }
                            else
                            {
                                // term not found, create term
                                if (curLevelTermObj.level == 0)
                                {
                                    // level 0 parent term is termset
                                    try
                                    {
                                        var newTerm = set.CreateTerm(curLevelTermObj.termName, lcid, curLevelTermObj.id);
                                        ctx.ExecuteQuery();
                                        tcout(" -- ", "term created");

                                        if (curLevelTermObj.labels.Any())
                                        {
                                            // add labels to new term
                                            foreach (var label in curLevelTermObj.labels)
                                            {
                                                newTerm.CreateLabel(label, lcid, false);
                                            }

                                            try
                                            {
                                                ctx.ExecuteQuery();
                                                tcout(" -- ", "label(s) created");
                                            }
                                            catch (Exception ex)
                                            {
                                                tcout(" -- ", "ERROR creating label(s)", GetExcMsg(ex));
                                            }
                                        }

                                        // load new term and add to collection
                                        try
                                        {
                                            ctx.Load(newTerm, a => a.Name, a => a.PathOfTerm);
                                            ctx.ExecuteQuery();
                                            lstMmdTerms.Add(newTerm);
                                        }
                                        catch (Exception ex)
                                        {
                                            tcout(" -- ", "ERROR loading new term", GetExcMsg(ex));
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        tcout(" -- ", "ERROR creating term", GetExcMsg(ex));
                                    }

                                }
                                else
                                {
                                    // any other level, backtrack to get parent from path, create term in parent
                                    var parentTermPath = curLevelTermObj.path.Substring(0, curLevelTermObj.path.LastIndexOf(';'));

                                    var parentTerm = lstMmdTerms.FirstOrDefault(x => GenUtil.MmdDenormalize(x.PathOfTerm).ToLower() == GenUtil.MmdDenormalize(parentTermPath).ToLower());

                                    if (parentTerm == null)
                                    {
                                        // this shouldn't happen, new terms are added to the collection for subsequent searches
                                        tcout(" -- ", "parent term not found, cannot create new term");
                                    }
                                    else
                                    {
                                        // parent term found, create term here
                                        try
                                        {
                                            var newTerm = parentTerm.CreateTerm(curLevelTermObj.termName, lcid, curLevelTermObj.id);
                                            ctx.ExecuteQuery();
                                            tcout(" -- ", "term created");

                                            if (curLevelTermObj.labels.Any())
                                            {
                                                // add labels to new term
                                                foreach (var label in curLevelTermObj.labels)
                                                {
                                                    newTerm.CreateLabel(label, lcid, false);
                                                }

                                                try
                                                {
                                                    ctx.ExecuteQuery();
                                                    tcout(" -- ", "label(s) created");
                                                }
                                                catch (Exception ex)
                                                {
                                                    tcout(" -- ", "ERROR creating label(s)", GetExcMsg(ex));
                                                }
                                            }

                                            // load new term and add to collection
                                            try
                                            {
                                                ctx.Load(newTerm, a => a.Name, a => a.PathOfTerm);
                                                ctx.ExecuteQuery();
                                                lstMmdTerms.Add(newTerm);
                                            }
                                            catch (Exception ex)
                                            {
                                                tcout(" -- ", "ERROR loading new term", GetExcMsg(ex));
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            tcout(" -- ", "ERROR creating term", GetExcMsg(ex));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void ImportTermsSimple()
        {
            var msg = "";

            tcout("SiteUrl", tbSiteUrl.Text);
            tcout("UserName", tbUsername.Text);
            tcout("Domain", tbDomain.Text);
            tcout("IsSPOnline", cbIsSPOnline.Checked);
            tcout("TermStore", tbTermStore.Text);
            tcout("TermGroup", tbTermGroup.Text);
            tcout("TermSet", tbTermSet.Text);
            tcout("TermStoreID", tbTermStoreID.Text);
            tcout("TermGroupID", tbTermGroupID.Text);
            tcout("TermSetID", tbTermSetID.Text);

            tcout("ImportType", rbImportTypeSimple.Checked ? "Simple" : "Advanced");
            tcout("SourceType", rbImportSourceText.Checked ? "Text" : (rbImportSourceExcel.Checked ? "Excel" : "SQL"));
            tcout("AppendNewLabelsToExistingTerms", cbAppendNewLabelsToExistingTerms.Checked);
            tcout("SourceFilePath", tbImportSourceFilePath.Text);
            tcout("TermSeparator", tbImportSeparator.Text);
            tcout("DbConnString", tbImportDbConnString.Text);
            tcout("SelectStmt", tbImportSelectStmt.Text);
            tcout("---------------------------------------------");

            var tmpSep = tbImportSeparator.Text.Trim() == "\\t" ? "\t" : tbImportSeparator.Text.Trim();

            // get terms and labels from source
            var lstTermObjs = new List<TermObj>();

            if (rbImportSourceText.Checked)
            {
                if (!ImportFileHelper.GetDataFromTextFileSimple(tmpSep, tbImportSourceFilePath.Text.Trim(), out lstTermObjs, out msg))
                {
                    tcout("ERROR extracting data from text file", msg);
                    return;
                }
            }
            else if (rbImportSourceExcel.Checked)
            {
                if (!ImportFileHelper.GetDataFromExcelFileSimple(tbImportSourceFilePath.Text.Trim(), out lstTermObjs, out msg))
                {
                    tcout("ERROR extracting data from excel file", msg);
                    return;
                }
            }
            else if (rbImportSourceSQL.Checked)
            {
                if (!ImportFileHelper.GetDataFromSqlSimple(tbImportDbConnString.Text.Trim(), tbImportSelectStmt.Text.Trim(), out lstTermObjs, out msg))
                {
                    tcout("ERROR extracting data from sql", msg);
                    return;
                }
            }

            if (!lstTermObjs.Any())
            {
                tcout("Found no terms to import");
                return;
            }

            var targetSite = new Uri(tbSiteUrl.Text.Trim());

            using (ClientContext ctx = new ClientContext(targetSite))
            {
                ctx.Credentials = BuildCreds();
                FixCtxForMixedMode(ctx);

                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQuery();
                tcout("Site loaded", web.Title);

                var lcid = System.Globalization.CultureInfo.CurrentCulture.LCID;
                var session = TaxonomySession.GetTaxonomySession(ctx);

                TermStore store = null;
                TermGroup group = null;
                TermSet set = null;

                LoadTStoreTGroupTSet(ctx, session, ref store, ref group, ref set);

                // import terms
                int i = 0;
                foreach (var termObj in lstTermObjs)
                {
                    i++;
                    var progress = i.ToString() + "/" + lstTermObjs.Count();
                    var curTermObj = termObj;

                    // find matching term in termset, if NOT found add term and labels
                    var matchInfo = new LabelMatchInformation(ctx)
                    {
                        TermLabel = curTermObj.termName,
                        TrimUnavailable = false
                    };
                    var termMatches = set.GetTerms(matchInfo);

                    try
                    {
                        ctx.Load(termMatches, x => x.Include(y => y.Labels));
                        ctx.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        tcout(progress, " *** ERROR searching for term matches", GetExcMsg(ex));
                        continue;
                    }

                    if (termMatches.Any())
                    {
                        // term found, optionally add new labels
                        var match = termMatches.First();

                        tcout(progress, "term found, exists", curTermObj.termName);

                        if (cbAppendNewLabelsToExistingTerms.Checked)
                        {
                            if (curTermObj.labels.Any())
                            {
                                foreach (var label in curTermObj.labels)
                                {
                                    var curLabel = label;

                                    var found = false;
                                    foreach (var matchLabel in match.Labels)
                                    {
                                        if (matchLabel.Value.IsEqual(curLabel))
                                        {
                                            found = true;
                                            break;
                                        }
                                    }

                                    if (!found)
                                    {
                                        // create label
                                        try
                                        {
                                            match.CreateLabel(curLabel, lcid, false);
                                            ctx.ExecuteQuery();
                                            tcout(progress, " -- term label added", curLabel);
                                        }
                                        catch (Exception ex)
                                        {
                                            tcout(progress, " *** ERROR creating new term label", curLabel, GetExcMsg(ex));
                                        }
                                    }
                                }
                            }
                        }

                    }
                    else
                    {
                        // term not found, creating new term and labels
                        tcout(progress, "term NOT found, creating", curTermObj.termName);

                        var newTerm = set.CreateTerm(curTermObj.termName, lcid, curTermObj.termId);

                        try
                        {
                            ctx.ExecuteQuery();
                            tcout(progress, " -- term created!");
                        }
                        catch (Exception ex)
                        {
                            tcout(progress, " *** ERROR creating new term", GetExcMsg(ex));
                            continue;
                        }

                        if (curTermObj.labels.Any())
                        {
                            foreach (var label in curTermObj.labels)
                            {
                                newTerm.CreateLabel(label, lcid, false);
                            }

                            try
                            {
                                ctx.ExecuteQuery();
                                tcout(progress, " -- term labels added!");
                            }
                            catch (Exception ex)
                            {
                                tcout(progress, " *** ERROR creating new term labels", GetExcMsg(ex));
                                continue;
                            }
                        }
                    }
                }
            }
        }

        private void bgw_btnStartImport_End(object sender, RunWorkerCompletedEventArgs e)
        {
            FlushCoutBuffer();
            SaveLogToFile("ImportTerms");
            EnableFormControls();
        }







        private void btnStartClearTermSet_Click(object sender, EventArgs e)
        {
            DialogResult dgResult = MessageBox.Show("Are you sure?", "Clear Term Set", MessageBoxButtons.YesNo);

            if (dgResult != DialogResult.Yes)
            {
                cout("Canceled");
                return;
            }

            DisableFormControls();
            InitCoutBuffer();
            tbStatus.Text = "";

            bgw = new BackgroundWorker();
            bgw.DoWork += new DoWorkEventHandler(bgw_btnStartClearTermSet);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_btnStartClearTermSet_End);
            bgw.ProgressChanged += new ProgressChangedEventHandler(BgwReportProgress);
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync();
        }

        private void bgw_btnStartClearTermSet(object sender, DoWorkEventArgs e)
        {
            try
            {
                ClearTermSet();

            }
            catch (Exception ex)
            {
                tcout("ERROR", GetExcMsg(ex));
            }
        }

        private void ClearTermSet()
        {
            tcout("SiteUrl", tbSiteUrl.Text);
            tcout("UserName", tbUsername.Text);
            tcout("Domain", tbDomain.Text);
            tcout("IsSPOnline", cbIsSPOnline.Checked);
            tcout("TermStore", tbTermStore.Text);
            tcout("TermGroup", tbTermGroup.Text);
            tcout("TermSet", tbTermSet.Text);
            tcout("TermStoreID", tbTermStoreID.Text);
            tcout("TermGroupID", tbTermGroupID.Text);
            tcout("TermSetID", tbTermSetID.Text);
            tcout("---------------------------------------------");

            var targetSite = new Uri(tbSiteUrl.Text.Trim());

            using (ClientContext ctx = new ClientContext(targetSite))
            {
                ctx.Credentials = BuildCreds();
                FixCtxForMixedMode(ctx);

                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQuery();
                tcout("Site loaded", web.Title);

                var lcid = System.Globalization.CultureInfo.CurrentCulture.LCID;
                var session = TaxonomySession.GetTaxonomySession(ctx);

                TermStore store = null;
                TermGroup group = null;
                TermSet set = null;

                LoadTStoreTGroupTSet(ctx, session, ref store, ref group, ref set);

                // get root terms for deletion
                var terms = set.Terms;

                ctx.Load(terms, a => a.Include(b => b.Name));
                ctx.ExecuteQuery();

                tcout("Root terms loaded", terms.Count);

                int i = 0;
                foreach (var curTerm in terms)
                {
                    i++;
                    var progress = i.ToString() + "/" + terms.Count.ToString();

                    tcout(progress, "Deleting term", curTerm.Name);
                    curTerm.DeleteObject();
                }
                ctx.ExecuteQuery();
            }
        }

        private void bgw_btnStartClearTermSet_End(object sender, RunWorkerCompletedEventArgs e)
        {
            FlushCoutBuffer();
            SaveLogToFile("ClearTermSet");
            EnableFormControls();
        }






        private void btnStartUpdate_Click(object sender, EventArgs e)
        {
            DialogResult dgResult = MessageBox.Show("Are you sure?", "Update Terms", MessageBoxButtons.YesNo);

            if (dgResult != DialogResult.Yes)
            {
                cout("Canceled");
                return;
            }

            DisableFormControls();
            InitCoutBuffer();
            tbStatus.Text = "";

            bgw = new BackgroundWorker();
            bgw.DoWork += new DoWorkEventHandler(bgw_btnStartUpdate);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_btnStartUpdate_End);
            bgw.ProgressChanged += new ProgressChangedEventHandler(BgwReportProgress);
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync();
        }

        private void bgw_btnStartUpdate(object sender, DoWorkEventArgs e)
        {
            try
            {
                UpdateTerms(); 
            }
            catch (Exception ex)
            {
                tcout("ERROR", GetExcMsg(ex));
            }
        }

        private void UpdateTerms()
        {
            string msg = "";
            List<SimpleImportObj> lstImportObjs = null;

            tcout("SiteUrl", tbSiteUrl.Text);
            tcout("UserName", tbUsername.Text);
            tcout("Domain", tbDomain.Text);
            tcout("IsSPOnline", cbIsSPOnline.Checked);
            tcout("TermStore", tbTermStore.Text);
            tcout("TermGroup", tbTermGroup.Text);
            tcout("TermSet", tbTermSet.Text);
            tcout("TermStoreID", tbTermStoreID.Text);
            tcout("TermGroupID", tbTermGroupID.Text);
            tcout("TermSetID", tbTermSetID.Text);
            tcout("UpdateTermsSourceFilePath", tbUpdateTermsSourceFilePath.Text);
            tcout("---------------------------------------------");

            if (!ImportFileHelper.GetUpdateSimpleDataFromExcelFile(tbUpdateTermsSourceFilePath.Text.Trim(), out lstImportObjs, out msg))
            {
                tcout("ERROR extracting data from Excel file", msg);
                return;
            }
            else if (!lstImportObjs.Any())
            {
                tcout("Nothing found to update.");
                return;
            }

            var targetSite = new Uri(tbSiteUrl.Text.Trim());

            using (ClientContext ctx = new ClientContext(targetSite))
            {
                ctx.Credentials = BuildCreds();
                FixCtxForMixedMode(ctx);

                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQuery();
                tcout("Site loaded", web.Title);

                var lcid = System.Globalization.CultureInfo.CurrentCulture.LCID;
                var session = TaxonomySession.GetTaxonomySession(ctx);

                TermStore store = null;
                TermGroup group = null;
                TermSet set = null;

                LoadTStoreTGroupTSet(ctx, session, ref store, ref group, ref set);

                // get all terms, then loop to find matches to update
                var terms = set.GetAllTerms();

                ctx.Load(terms, a =>
                    a.Include(
                        b => b.Id,
                        b => b.Name,
                        b => b.PathOfTerm,
                        b => b.Description,
                        b => b.IsAvailableForTagging,
                        b => b.IsReused,
                        b => b.IsSourceTerm,
                        b => b.Labels.Include(
                            c => c.IsDefaultForLanguage,
                            c => c.Value),
                        b => b.Parent.Id));
                ctx.ExecuteQuery();

                tcout("All terms loaded", terms.Count);

                if (!terms.Any())
                {
                    tcout("No terms found in MMD to update.");
                    return;
                }
                else
                {
                    foreach (var importObj in lstImportObjs)
                    {
                        foreach (var curTerm in terms)
                        {
                            // matching using term ids only (not term names!)
                            if (curTerm.Id == importObj.termId)
                            {
                                if (importObj.termName.IsEqual("$delete") || importObj.termName.IsEqual("#delete")) 
                                { 
                                    // delete term
                                    tcout("Match found, deleting term", curTerm.Id, curTerm.Name);

                                    try
                                    {
                                        curTerm.DeleteObject();
                                        ctx.ExecuteQuery();
                                        tcout(" -- Term deleted");
                                    }
                                    catch (Exception ex)
                                    {
                                        tcout(" *** ERROR deleting term", GetExcMsg(ex));
                                    }
                                }
                                else
                                {
                                    // update term
                                    tcout("Match found, updating term", curTerm.Id, curTerm.Name);

                                    try
                                    {
                                        // make sure new termname is not the current termname, and not a current label
                                        if (!GenUtil.MmdDenormalize(curTerm.Name).IsEqual(GenUtil.MmdDenormalize(importObj.termName)))
                                        {
                                            if (curTerm.Labels.Any(x => !x.IsDefaultForLanguage && GenUtil.MmdDenormalize(x.Value).ToLower() == GenUtil.MmdDenormalize(importObj.termName).ToLower()))
                                            {
                                                tcout(" -- Cannot update Term Name, matches existing label");
                                            }
                                            else
                                            {
                                                curTerm.Name = GenUtil.MmdDenormalize(importObj.termName);
                                                ctx.ExecuteQuery();
                                                tcout(" -- Term Name updated");
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        tcout(" *** ERROR updating Term Name", GetExcMsg(ex));
                                    }

                                    if (!curTerm.Description.IsEqual(importObj.descr))
                                    {
                                        try
                                        {
                                            curTerm.SetDescription(importObj.descr, System.Globalization.CultureInfo.CurrentCulture.LCID);
                                            ctx.ExecuteQuery();
                                            tcout(" -- Description updated");
                                        }
                                        catch (Exception ex)
                                        {
                                            tcout(" *** ERROR updating Description", GetExcMsg(ex));
                                        }
                                    }

                                    if (curTerm.IsAvailableForTagging != importObj.isAvailForTagging)
                                    {
                                        try
                                        {
                                            curTerm.IsAvailableForTagging = importObj.isAvailForTagging;
                                            ctx.ExecuteQuery();
                                            tcout(" -- IsAvailableForTagging updated");
                                        }
                                        catch (Exception ex)
                                        {
                                            tcout(" *** ERROR updating IsAvailableForTagging", GetExcMsg(ex));
                                        }
                                    }

                                    if (!cbSkipUpdatingTermLabels.Checked)
                                    {
                                        try
                                        {
                                            // add new labels found in new collection and not in term.labels
                                            var labelsToAdd = new List<string>();
                                            foreach (var lbl in importObj.labels)
                                            {
                                                var curLbl = lbl;
                                                if (!curTerm.Labels.Any(x => !x.IsDefaultForLanguage && GenUtil.MmdDenormalize(x.Value).ToLower() == GenUtil.MmdDenormalize(curLbl).ToLower()))
                                                {
                                                    labelsToAdd.Add(curLbl);
                                                }
                                            }

                                            // remove old labels not in new collection
                                            var labelsToRemove = new List<string>();
                                            foreach (var lbl in curTerm.Labels)
                                            {
                                                var curLbl = lbl;
                                                if (!importObj.labels.Any(x => GenUtil.MmdDenormalize(x).ToLower() == GenUtil.MmdDenormalize(curLbl.Value).ToLower()))
                                                {
                                                    if (!curLbl.IsDefaultForLanguage)
                                                    {
                                                        labelsToRemove.Add(curLbl.Value);
                                                    }
                                                }
                                            }

                                            foreach (var lbl in labelsToAdd)
                                            {
                                                curTerm.CreateLabel(GenUtil.MmdDenormalize(lbl), System.Globalization.CultureInfo.CurrentCulture.LCID, false);
                                                ctx.ExecuteQuery();
                                                tcout(" -- Term label added", lbl);
                                            }

                                            foreach (var lbl in labelsToRemove)
                                            {
                                                curTerm.Labels.GetByValue(lbl).DeleteObject(); // this works when label has '&' in it
                                                ctx.ExecuteQuery();
                                                tcout(" -- Term label deleted", lbl);
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            tcout(" *** ERROR updating syncing Term Labels", GetExcMsg(ex));
                                        }
                                    }
                                }

                                break;
                            }
                        }
                    }

                }

                
            }
        }

        private void bgw_btnStartUpdate_End(object sender, RunWorkerCompletedEventArgs e)
        {
            FlushCoutBuffer();
            SaveLogToFile("UpdateTerms");
            EnableFormControls();
        }












        private void btnLoadMMD_Click(object sender, EventArgs e)
        {
            DisableFormControls();
            InitCoutBuffer();
            tbStatus.Text = "";

            bgw = new BackgroundWorker();
            bgw.DoWork += new DoWorkEventHandler(bgw_btnLoadMMD);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_btnLoadMMD_End);
            bgw.ProgressChanged += new ProgressChangedEventHandler(BgwReportProgress);
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync();
        }

        private void bgw_btnLoadMMD(object sender, DoWorkEventArgs e)
        {
            TreeNode rootNode = null;

            try
            {
                rootNode = new TreeNode("MMD");
                rootNode.Tag = "-1";

                var targetSite = new Uri(tbSiteUrl.Text.Trim());

                using (ClientContext ctx = new ClientContext(targetSite))
                {
                    ctx.Credentials = BuildCreds();
                    FixCtxForMixedMode(ctx);

                    Web web = ctx.Web;
                    ctx.Load(web, w => w.Title);
                    ctx.ExecuteQuery();
                    tcout("Site loaded", web.Title);

                    var lcid = System.Globalization.CultureInfo.CurrentCulture.LCID;
                    var session = TaxonomySession.GetTaxonomySession(ctx);

                    // load all termstores available
                    var termStores = session.TermStores;
                    ctx.Load(termStores, x => x.Include(y => y.Id, y => y.Name));
                    ctx.ExecuteQuery();

                    tcout(" - found # termstores", termStores.Count);

                    foreach (var termStore in termStores)
                    {
                        var termStoreNode = rootNode.Nodes.Add(termStore.Name);
                        termStoreNode.Tag = termStore.Id.ToString();

                        // load all term groups
                        var groups = termStore.Groups;
                        ctx.Load(groups, x => x.Include(y => y.Name, y => y.Id));
                        ctx.ExecuteQuery();

                        tcout(" -- found # termgroups", groups.Count);

                        foreach (var group in groups)
                        {
                            var groupNode = termStoreNode.Nodes.Add(group.Name);
                            groupNode.Tag = group.Id.ToString();

                            // load all termsets
                            var termSets = group.TermSets;
                            ctx.Load(termSets, x => x.Include(y => y.Name, y => y.Id));
                            ctx.ExecuteQuery();

                            tcout(" --- found # termsets", termSets.Count);

                            foreach (var termSet in termSets)
                            {
                                var termSetNode = groupNode.Nodes.Add(termSet.Name);
                                termSetNode.Tag = termSet.Id.ToString();

                                if (cbShowEmptyTermSetsWithColor.Checked)
                                {
                                    var terms = termSet.Terms;
                                    ctx.Load(terms, x => x.Include(y => y.Id));
                                    ctx.ExecuteQuery();

                                    var termsCount = termSet.Terms.Count;

                                    if (termsCount > 0)
                                    {
                                        termSetNode.ForeColor = System.Drawing.Color.Black;
                                    }
                                    else
                                    {
                                        termSetNode.ForeColor = System.Drawing.Color.LightGray;
                                    }
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                tcout("ERROR", GetExcMsg(ex));
            }

            e.Result = new List<object>() { rootNode };
        }

        private void bgw_btnLoadMMD_End(object sender, RunWorkerCompletedEventArgs e)
        {
            var lst = e.Result as List<object>;
            var rootNode = lst[0] as TreeNode;

            tvMMD.Nodes.Clear();
            tvMMD.Nodes.Add(rootNode);


            FlushCoutBuffer();
            SaveLogToFile("LoadMMD");
            EnableFormControls();
        }

        private void tvMMD_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Level == 1)
            {
                // termstore
                tbTermStore.Text = e.Node.Text;
                tbTermStoreID.Text = e.Node.Tag.ToString();

                tbTermGroup.Text = "";
                tbTermGroupID.Text = "";

                tbTermSet.Text = "";
                tbTermSetID.Text = "";
            }
            else if (e.Node.Level == 2)
            {
                // group
                tbTermStore.Text = e.Node.Parent.Text;
                tbTermStoreID.Text = e.Node.Parent.Tag.ToString();

                tbTermGroup.Text = e.Node.Text;
                tbTermGroupID.Text = e.Node.Tag.ToString();

                tbTermSet.Text = "";
                tbTermSetID.Text = "";
            }
            else if (e.Node.Level == 3)
            {
                // termset
                tbTermStore.Text = e.Node.Parent.Parent.Text;
                tbTermStoreID.Text = e.Node.Parent.Parent.Tag.ToString();

                tbTermGroup.Text = e.Node.Parent.Text;
                tbTermGroupID.Text = e.Node.Parent.Tag.ToString();

                tbTermSet.Text = e.Node.Text;
                tbTermSetID.Text = e.Node.Tag.ToString();
            }
            else
            {
                // clear all
                tbTermStore.Text = "";
                tbTermStoreID.Text = "";

                tbTermGroup.Text = "";
                tbTermGroupID.Text = "";

                tbTermSet.Text = "";
                tbTermSetID.Text = "";
            }
        }











        /// <summary>
        /// Combine function params as strings with separator, no line breaks.
        /// </summary>
        public string CombineFnParmsToString(params object[] objs)
        {
            string output = "";
            string delim = ": ";

            for (int i = 0; i < objs.Length; i++)
            {
                if (objs[i] == null) objs[i] = "";
                if (i == objs.Length - 1) delim = "";
                output += string.Concat(objs[i], delim);
            }

            return output;
        }

        /// <summary>
        /// Build message for status window, prepend datetime, append message (already combined with separator), append newline chars.
        /// </summary>
        public string BuildCoutMessage(string msg)
        {
            return string.Format("{0}: {1}\r\n", DateTime.Now.ToLongTimeString(), msg);
        }

        /// <summary>
        /// Standard status dumping function, immediately dumps to status window.
        /// </summary>
        public void cout(params object[] objs)
        {
            tbStatus.AppendText(BuildCoutMessage(CombineFnParmsToString(objs)));
        }

        string tcout_buffer;
        int tcout_counter;

        /// <summary>
        /// Threaded status dumping function, uses buffer to only dump to status window peridocially, batch size configured in app.config.
        /// </summary>
        public void tcout(params object[] objs)
        {
            tcout_counter++;
            tcout_buffer += BuildCoutMessage(CombineFnParmsToString(objs));

            var batchSize = statusWindowOutputBatchSize == 0 ? 1 : statusWindowOutputBatchSize;

            if (tcout_counter % batchSize == 0)
            {
                bgw.ReportProgress(0, tcout_buffer);
                InitCoutBuffer();
            }
        }

        /// <summary>
        /// Reset status buffer.
        /// </summary>
        private void InitCoutBuffer()
        {
            tcout_counter = 0;
            tcout_buffer = "";
        }

        /// <summary>
        /// Flush status buffer to status window (since using mod operator).
        /// </summary>
        private void FlushCoutBuffer()
        {
            if (!tcout_buffer.IsNull())
            {
                tbStatus.AppendText(tcout_buffer);
                InitCoutBuffer();
            }
        }

        /// <summary>
        /// Threaded callback function, dump input to status window, already formatted with datetime, combined params, and linebreaks.
        /// </summary>
        private void BgwReportProgress(object sender, ProgressChangedEventArgs e)
        {
            tbStatus.AppendText(e.UserState.ToString());
        }





        private void lnkClear_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tbStatus.ResetText();
        }

        private void lnkExport_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveLogToFile(null);
            MessageBox.Show("Log saved to EXE folder.");
        }

        void SaveLogToFile(string action)
        {
            if (!action.IsNull())
            {
                action = action.Trim().ToUpper() + "_";
            }

            var exportFilePath = AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists(exportFilePath.CombineFS("logs")))
                Directory.CreateDirectory(exportFilePath.CombineFS("logs"));
            exportFilePath = exportFilePath.CombineFS("logs\\log" + "_" + action.SafeTrim() + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".txt");

            System.IO.File.WriteAllText(exportFilePath, tbStatus.Text + "\r\n[EOF]");

            cout("Log saved to EXE folder.");
        }







        private string GetExcMsg(Exception ex)
        {
            if (showFullErrMsgs)
                return ex.ToString();
            else
                return ex.Message;
        }










        private void rbExportTypeSimple_CheckedChanged(object sender, EventArgs e)
        {
            SetExportVis();
        }

        private void rbExportTypeAdv_CheckedChanged(object sender, EventArgs e)
        {
            SetExportVis();
        }

        private void SetExportVis()
        {
            if (rbExportTypeSimple.Checked)
            {
                cbExportTermIds.Enabled = false;
                cbExportTermLabels.Enabled = false;
            }
            else
            {
                cbExportTermIds.Enabled = true;
                cbExportTermLabels.Enabled = true;
            }
        }









        private void rbImportSourceText_CheckedChanged(object sender, EventArgs e)
        {
            SetImportVis();
        }

        private void rbImportSourceExcel_CheckedChanged(object sender, EventArgs e)
        {
            SetImportVis();
        }

        private void rbImportSourceSQL_CheckedChanged(object sender, EventArgs e)
        {
            SetImportVis();
        }

        private void SetImportVis()
        {
            if (rbImportSourceText.Checked)
            {
                tbImportSourceFilePath.Enabled = true;
                tbImportSeparator.Enabled = true;
                tbImportDbConnString.Enabled = false;
                tbImportSelectStmt.Enabled = false;
            }
            else if (rbImportSourceExcel.Checked)
            {
                tbImportSourceFilePath.Enabled = true;
                tbImportSeparator.Enabled = false;
                tbImportDbConnString.Enabled = false;
                tbImportSelectStmt.Enabled = false;
            }
            else
            {
                tbImportSourceFilePath.Enabled = false;
                tbImportSeparator.Enabled = false;
                tbImportDbConnString.Enabled = rbImportTypeSimple.Checked ? true : false;
                tbImportSelectStmt.Enabled = rbImportTypeSimple.Checked ? true : false;
            }
        }






        private void rbImportSimple_CheckedChanged(object sender, EventArgs e)
        {
            SetImportTypeVis();
        }

        private void rbImportAdvanced_CheckedChanged(object sender, EventArgs e)
        {
            SetImportTypeVis();
        }

        void SetImportTypeVis()
        {
            if (rbImportTypeSimple.Checked)
            {
                rbImportSourceSQL.Enabled = true;
            }
            else
            {
                rbImportSourceSQL.Enabled = false;

                if (rbImportSourceSQL.Checked)
                {
                    rbImportSourceSQL.Checked = false;
                    rbImportSourceText.Checked = true;
                }
            }

            SetImportVis();
        }

        private void rbImportTypeSimple_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Imports terms from file, terms are flat, supports term name and optionally labels (specify separator if source is TEXT).";
        }

        private void rbImportTypeAdvanced_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Imports terms from file, terms can be hierarchical, supports Id, term name, labels. See sample file for import format.";
        }

        private void rbExportTypeSimple_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Exports all terms in termset flat, with extra info, good for DB importing or Excel analysis.";
        }

        private void rbExportTypeAdv_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Exports all terms in hierarchical format, preserving parent/child relationships, good for migrating Termsets in SP.";
        }

        private void rbExportTypeSimple_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void rbExportTypeAdv_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void rbImportTypeSimple_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void rbImportTypeAdvanced_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void tbExportFilePath_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Enter path to directory where file will be created, or leave blank for current EXE path.";
        }

        private void tbExportFilePath_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void tbUpdateTermsSourceFilePath_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Enter path to Excel file where terms were exported into Simple Format and changes needs to be pushed to MMD.";
        }

        private void tbUpdateTermsSourceFilePath_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void imageBandR_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.bandrsolutions.com/?utm_source=SPTaxonomyToolsOnline&utm_medium=application&utm_campaign=SPTaxonomyToolsOnline");
        }

        private void imageBandRwait_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.bandrsolutions.com/?utm_source=SPTaxonomyToolsOnline&utm_medium=application&utm_campaign=SPTaxonomyToolsOnline");
        }






    }
}
