using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using EMecaAddin.Media;
using EMecaAddin.Properties;
using E_Meca.TypeSpeed;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace EMecaAddin
{
    public partial class Ribbon
    {
        private ICharacterCost _characterCost;
        private ILevenshteinDistance _levenshteinDistance;
        private readonly ISound _sound = new Sound();

        private string _model = "";
        private long _testTimeInMilliseconds;
        private readonly Stopwatch _stopwatch = new Stopwatch();
        private static readonly int _incorrectKeyPenalization = 1;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            _characterCost = new CharacterCost();
            _levenshteinDistance = new LevenshteinDistance(_characterCost);
            LoadTimeOptions();
        }

        private void LoadTimeOptions()
        {
            ddTime.Items.Clear();

            var firstItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            firstItem.Label = Resources.NoTime;
            firstItem.Tag = 0;
            ddTime.Items.Add(firstItem);

            for (var i = 1; i <= 20; i++)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = $@"{i} {(i == 1 ? Resources.Minute : Resources.Minutes)}";
                item.Tag = i;
                ddTime.Items.Add(item);
            }
        }

        private void BtnGetTypingData_Click(object sender, RibbonControlEventArgs e)
        {
            ShowTypingData();
        }

        private void ShowTypingData()
        {
            var documentText = GetDocumentText();
            var keysPressed = _characterCost.ComputeText(documentText);
            var incorrectKeysPressed = ComputeErrors(_model, documentText);
            var correctKeysPressed = keysPressed - incorrectKeysPressed;
            var charactersCount = documentText.Length;

            ShowData(charactersCount, keysPressed, correctKeysPressed, _testTimeInMilliseconds, incorrectKeysPressed);
        }

        private int ComputeErrors(string model, string documentText)
        {
            return _levenshteinDistance.Compute(model, documentText);
        }

        private static void ShowData(int charactersCount, int keysPressed, int correctKeysPressed, long testTimeInMilliseconds, int incorrectKeysPressed)
        {
            const string separator = ": ";
            var sb = new StringBuilder();

            sb.Append(Resources.Characters).Append(separator).AppendLine(charactersCount.ToString(CultureInfo.CurrentCulture));
            sb.Append(Resources.KeysPressed).Append(separator).AppendLine(keysPressed.ToString(CultureInfo.CurrentCulture));
            if (testTimeInMilliseconds != 0)
            {
                var minutes = testTimeInMilliseconds / 60000m;
                var resultantKeysPressed = correctKeysPressed - incorrectKeysPressed * _incorrectKeyPenalization;
                var typeSpeed = Math.Round(resultantKeysPressed / minutes, MidpointRounding.AwayFromZero);
                var time = TimeSpan.FromMilliseconds(testTimeInMilliseconds);

                sb.Append(Resources.CorrectKeysPressed).Append(separator).AppendLine(correctKeysPressed.ToString(CultureInfo.CurrentCulture));
                sb.Append(Resources.IncorrectKeysPressed).Append(separator).AppendLine(incorrectKeysPressed.ToString(CultureInfo.CurrentCulture));
                sb.Append(Resources.Time).Append(separator).AppendLine(time.ToString(@"hh\:mm\:ss\.f"));
                sb.AppendLine();
                sb.Append(Resources.TypeSpeed).Append(separator).AppendLine(typeSpeed.ToString(CultureInfo.CurrentCulture));
            }

            MessageBox.Show(sb.ToString());
        }

        private static string GetDocumentText()
        {
            return Globals.ThisAddIn.Application.ActiveDocument.Content.Text.TrimEnd('\r');
        }

        private void BtnOpenModel_Click(object sender, RibbonControlEventArgs e)
        {
            var dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
                _model = GetTextFromWord(openFileDialog.FileName);
        }

        private string GetTextFromWord(string filePath)
        {
            var word = new Microsoft.Office.Interop.Word.Application();
            object miss = Missing.Value;
            object path = filePath;
            object readOnly = true;
            var docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            docs.Content.Select();
            return word.Application.Selection.Text;
        }

        private void BtnStart_Click(object sender, RibbonControlEventArgs e)
        {
            StartTest();
        }

        private void StartTest()
        {
            if (string.IsNullOrEmpty(_model))
            {
                var response = MessageBox.Show(Resources.NoModelMessage, Resources.Warning, MessageBoxButtons.OKCancel, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button2);
                if (response == DialogResult.Cancel) return;
            }

            PrepareDocumentForNewTest();

            btnStart.Visible = false;
            btnStop.Visible = true;

            _testTimeInMilliseconds = GetSelectedTimeInMilliseconds();

            if (_stopwatch.IsRunning) _stopwatch.Reset();
            _stopwatch.Start();

            if (_testTimeInMilliseconds > 0)
                timerTest.Start();
        }

        private void PrepareDocumentForNewTest()
        {
            if (Globals.ThisAddIn.Application.ActiveDocument.ProtectionType != WdProtectionType.wdNoProtection)
                Globals.ThisAddIn.Application.ActiveDocument.Unprotect("");

            Globals.ThisAddIn.Application.ActiveDocument.Content.Delete();
        }

        private int GetSelectedTimeInMilliseconds()
        {
            return Convert.ToInt32(Math.Round(Convert.ToDecimal(ddTime.SelectedItem.Tag) * 60000));
        }

        private void BtnStop_Click(object sender, RibbonControlEventArgs e)
        {
            StopTest();
        }

        private void StopTest()
        {
            _stopwatch.Stop();
            _testTimeInMilliseconds = _stopwatch.ElapsedMilliseconds;
            _stopwatch.Reset();
            timerTest.Stop();
            PlaySound();
            Globals.ThisAddIn.Application.ActiveDocument.Protect(WdProtectionType.wdAllowOnlyReading, false, "", false, false);
            btnStart.Visible = true;
            btnStop.Visible = false;
            ShowTypingData();
        }

        private void TimerTest_Tick(object sender, EventArgs e)
        {
            if (_stopwatch.ElapsedMilliseconds >= _testTimeInMilliseconds)
                StopTest();
        }

        private void btnNewTest_Click(object sender, RibbonControlEventArgs e)
        {
            PrepareDocumentForNewTest();
        }

        private void PlaySound()
        {
            _sound.Play();
        }
    }
}
