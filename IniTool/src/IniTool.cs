using System;
using System.Text;

using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Text.RegularExpressions;

namespace Tools.IO
{
    /// <summary>
    /// Used for writing and reading ini-files. 
    /// </summary>
    public class IniTool
    {
        /*****************************************************************/
        // Declarations
        /*****************************************************************/
        #region Declarations

        /// <summary>
        /// File path that leads to the ini file. 
        /// </summary>
        public string FilepathIni { get; set; }

        // Imports for reading and writing to an ini file. 
        [DllImport("kernel32", EntryPoint = "WritePrivateProfileString")]
        static extern int WriteProfile(string cSection, string cEntry, string cValue, string fileName);
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileString")]
        static extern int ReadProfile(string cSection, string cEntry, string cValue, StringBuilder result, int size, string fileName);
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileSection")]
        static extern int ReadSection(string section, byte[] returnBytes, int size, string fileName);
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileSectionNamesA")]
        static extern int GetPrivateProfileSectionNames(byte[] lpszReturnBuffer, int nSize, string lpFileName);

        #endregion Declarations
        /*****************************************************************/
        // Constructors
        /*****************************************************************/
        #region Constructors

        /// <summary>
        /// Creates an instance of this class and determines the ini file path via the exe-file path of the application. 
        /// </summary>
        public IniTool()
        {
            string sExeDir = Path.GetDirectoryName(Application.ExecutablePath);
            string sExeName = Path.GetFileNameWithoutExtension(Application.ExecutablePath);

            this.FilepathIni = sExeDir + "\\" + sExeName + ".ini";
        }

        /// <summary>
        /// Creates an instance of this class and uses the given ini file path. 
        /// </summary>
        /// <param name="filepathIni"></param>
        public IniTool(string filepathIni)
        {
            this.FilepathIni = filepathIni;
        }

        #endregion Constructors
        /*****************************************************************/
        // Methods
        /*****************************************************************/
        #region Methods

        #region Exists

        /// <summary>
        /// Returns true, if the given ini file contains an entity 
        /// with the given identifier in the given section. 
        /// </summary>
        /// <param name="sSection"></param>
        /// <param name="sEntry"></param>
        /// <returns></returns>
        public bool Exists(string sSection, string sEntry)
        {
            string sEntryExists = this.Get(sSection, sEntry);

            if (string.IsNullOrWhiteSpace(sEntryExists))
                return false;
            else
                return true;
        }

        #endregion Exists

        #region Get

        /// <summary>
        /// Returns the value of a given entry in the given section. 
        /// </summary>
        /// <param name="sSection">Section the entry is in. </param>
        /// <param name="sEntry">Entry to get the value from. </param>
        /// <returns></returns>
        public string Get(string sSection, string sEntry)
        {
            StringBuilder sbValue = new StringBuilder(4096);

            try
            {
                ReadProfile(sSection, sEntry, "", sbValue, sbValue.Capacity, this.FilepathIni);
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error reading entry: " + exc.Message);
                return null;
            }

            string sValue = sbValue.ToString();

            if (string.IsNullOrWhiteSpace(sValue)) // Entry does not exist. 
                return null;
            else
                return sValue;
        }

        /// <summary>
        /// Returns the value of a given entry in the given section. 
        /// If the given entry does not yet exist it will be created and the default value returned. 
        /// </summary>
        /// <param name="sSection">Section the entry is in. </param>
        /// <param name="sEntry">Entry to get the value from. </param>
        /// <param name="sDefaultValue">A default value to assign the entry if it doesn't yet exist. </param>
        /// <returns></returns>
        public string Get(string sSection, string sEntry, string sDefaultValue)
        {
            StringBuilder sbValue = new StringBuilder(4096);

            try
            {
                ReadProfile(sSection, sEntry, "", sbValue, sbValue.Capacity, this.FilepathIni);
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error reading entry: " + exc.Message);
                return null;
            }

            string sValue = sbValue.ToString();

            if (string.IsNullOrWhiteSpace(sValue)) // Entry does not yet exist. 
            {
                this.Set(sSection, sEntry, sDefaultValue);

                return sDefaultValue;
            }

            return sValue;
        }

        /// <summary>
        /// Returns an array of all entries of a given section. 
        /// </summary>
        /// <param name="sSection"></param>
        /// <returns></returns>
        public string[] GetAllEntries(string sSection)
        {
            byte[] arrBuffer = new byte[65536];

            try
            {
                ReadSection(sSection, arrBuffer, arrBuffer.Length, this.FilepathIni);
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error reading entry: " + exc.Message);
                return null;
            }

            // Contains all entries of the section. 
            string sEntries = Encoding.Default.GetString(arrBuffer);

            // An array of all individual entry names. 
            string[] arrEntry = sEntries.Split(new string[] { "\0" }, StringSplitOptions.RemoveEmptyEntries);

            return arrEntry;
        }

        /// <summary>
        /// Returns an array of all sections of the ini file. 
        /// </summary>
        /// <returns></returns>
        public string[] GetAllSections()
        {
            byte[] arrBuffer = new byte[65536];

            try
            {
                GetPrivateProfileSectionNames(arrBuffer, arrBuffer.Length, this.FilepathIni);
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error reading entry: " + exc.Message);
                return null;
            }

            // Contains all sections. 
            string sSections = Encoding.Default.GetString(arrBuffer);

            // Array of individual section names. 
            string[] arrSectionName = sSections.Split(new string[] { "\0" }, StringSplitOptions.RemoveEmptyEntries);

            return arrSectionName;
        }

        #endregion Get

        #region Set

        /// <summary>
        /// Adds or overrides the value of the given entry in the given section. 
        /// </summary>
        /// <param name="sSection">The section to set the entry for. </param>
        /// <param name="sEntry">The entry whose value to set. </param>
        /// <param name="sValue">The value to set. </param>
        public void Set(string sSection, string sEntry, string sValue)
        {
            try
            {
                WriteProfile(sSection, sEntry, sValue, this.FilepathIni);
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error writing entry: " + exc.Message);
                return;
            }
        }

        #endregion Set

        #region Remove

        /// <summary>
        /// Removes the given section. 
        /// </summary>
        /// <param name="sSection">Name of the section to remove. </param>
        public void Remove(string sSection)
        {
            try
            {
                WriteProfile(sSection, null, null, this.FilepathIni);
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error writing entry: " + exc.Message);
                return;
            }
        }

        /// <summary>
        /// Removes the given entry from the given section. 
        /// </summary>
        /// <param name="sSection">The section from which to remove the given entry. </param>
        /// <param name="sEntry">Name of the entry to remove. </param>
        public void Remove(string sSection, string sEntry)
        {
            try
            {
                WriteProfile(sSection, sEntry, null, this.FilepathIni);
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error writing entry: " + exc.Message);
                return;
            }
        }

        #endregion Remove

        #region FormLocation

        /// <summary>
        /// Writes the Location-property of the given form to the given ini file. 
        /// </summary>
        /// <param name="oForm"></param>
        public void WriteFormLocation(Form oForm)
        {
            this.Set("UI", oForm.Name + "Location", oForm.Location.ToString());
        }

        /// <summary>
        /// Reads the Location-property of the given form from the given ini file. 
        /// Overrides the Location-property of the given form. 
        /// </summary>
        /// <param name="oForm"></param>
        /// <returns></returns>
        public Point ReadFormLocation(Form oForm)
        {
            string sValue = this.Get("UI", oForm.Name + "Location", oForm.Location.ToString());

            Regex rgxValue = new Regex(@"\{[^=]+=(?<X>[0-9]+),[^=]+=(?<Y>[0-9]+)}", RegexOptions.Singleline);
            Match mValue = rgxValue.Match(sValue);

            int iX = oForm.Location.X;
            int iY = oForm.Location.Y;

            if (mValue.Groups["X"].Success)
                iX = int.Parse(mValue.Groups["X"].Value);

            if (mValue.Groups["Y"].Success)
                iY = int.Parse(mValue.Groups["Y"].Value);

            oForm.Location = new Point(iX, iY);

            return new Point(iX, iY);
        }

        #endregion FormLocation

        #region FormSize

        /// <summary>
        /// Writes the Size-property of the given form to the given ini file. 
        /// </summary>
        /// <param name="oForm"></param>
        public void WriteFormSize(Form oForm)
        {
            this.Set("UI", oForm.Name + "Size", oForm.Size.ToString());
        }

        /// <summary>
        /// Reads the Size-property of the given form from the given ini file. 
        /// Overrides the Location-property of the given form. 
        /// </summary>
        /// <param name="oForm"></param>
        /// <returns></returns>
        public Size ReadFormSize(Form oForm)
        {
            string sValue = this.Get("UI", oForm.Name + "Size", oForm.Size.ToString());

            Regex rgxValue = new Regex(@"\{[^=]+=(?<W>[0-9]+),[^=]+=(?<H>[0-9]+)}", RegexOptions.Singleline);
            Match mValue = rgxValue.Match(sValue);

            int iW = oForm.Size.Width;
            int iH = oForm.Size.Height;

            if (mValue.Groups["W"].Success)
                iW = int.Parse(mValue.Groups["W"].Value);

            if (mValue.Groups["H"].Success)
                iH = int.Parse(mValue.Groups["H"].Value);

            oForm.Size = new Size(iW, iH);

            return new Size(iW, iH);
        }

        #endregion FormSize

        #region Color

        /// <summary>
        /// Writes the given color value to the given entry. 
        /// </summary>
        /// <param name="sSection"></param>
        /// <param name="sEntry"></param>
        /// <param name="Color"></param>
        public void SetColor(string sSection, string sEntry, Color Color)
        {
            StringBuilder sbColor = new StringBuilder(20); // Initialize with sufficient capacity. 
            sbColor.Append("{");
            sbColor.Append("A");
            sbColor.Append(Color.A.ToString());
            sbColor.Append(";R");
            sbColor.Append(Color.R.ToString());
            sbColor.Append(";G");
            sbColor.Append(Color.G.ToString());
            sbColor.Append(";B");
            sbColor.Append(Color.B.ToString());
            sbColor.Append("}");

            this.Set(sSection, sEntry, sbColor.ToString());
        }

        /// <summary>
        /// Returns a color parsed from the value of the given entry. 
        /// </summary>
        /// <param name="sSection"></param>
        /// <param name="sEntry"></param>
        /// <returns></returns>
        public Color GetColor(string sSection, string sEntry)
        {
            string sValueEntry = this.Get(sSection, sEntry);

            // Finds a single letter followed by one or more decimal numbers. 
            string sPattern = @"(?<channel>[aArRgGbB])(?<value>\d+)";
            Regex rgxValues = new Regex(sPattern, RegexOptions.Singleline);

            MatchCollection mcValue = rgxValues.Matches(sValueEntry);

            int A = 0;
            int R = 0;
            int G = 0;
            int B = 0;

            foreach (Match mValue in mcValue)
            {
                string sChannel = mValue.Groups["channel"].Value.ToUpper();
                string sValue = mValue.Groups["value"].Value;

                if (sChannel == "A")
                    int.TryParse(sValue, out A);
                else if (sChannel == "R")
                    int.TryParse(sValue, out R);
                else if (sChannel == "G")
                    int.TryParse(sValue, out G);
                else if (sChannel == "B")
                    int.TryParse(sValue, out B);
            }

            return Color.FromArgb(A, R, G, B);
        }

        /// <summary>
        /// Returns a color parsed from the value of the given entry. 
        /// If the entry does not yet exist it will be created. 
        /// </summary>
        /// <param name="sSection"></param>
        /// <param name="sEntry"></param>
        /// <param name="ColorDefault"></param>
        /// <returns></returns>
        public Color GetColor(string sSection, string sEntry, Color ColorDefault)
        {
            string sValueEntry = this.Get(sSection, sEntry);

            if (string.IsNullOrWhiteSpace(sValueEntry))
            {
                this.SetColor(sSection, sEntry, ColorDefault);

                return ColorDefault;
            }
            else
            {
                return this.GetColor(sSection, sEntry);
            }
        }

        #endregion Color

        /// <summary>
        /// Shows a form through which the user can edit the ini entries. 
        /// </summary>
        public void ShowDialog()
        {
            Form frmIniSettings = new Form();
            frmIniSettings.Text = "Edit Ini File Entries";
            frmIniSettings.StartPosition = FormStartPosition.Manual;
            frmIniSettings.Size = new Size(500, 350);
            frmIniSettings.SizeGripStyle = SizeGripStyle.Show;
            frmIniSettings.KeyPreview = true;
            frmIniSettings.KeyDown += FrmIniSettings_KeyDown;

            // Distance between controls. 
            int pxMargin = 4;

            // Panel that will contain the controls displaying the INI-settings. 
            Panel pnl = new Panel();
            pnl.AutoScroll = true;
            pnl.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            pnl.Size = new Size(
                frmIniSettings.Width - (SystemInformation.VerticalScrollBarWidth + pxMargin),
                frmIniSettings.Height - (SystemInformation.HorizontalScrollBarHeight + pxMargin)
            );
            pnl.Location = new Point(pxMargin, pxMargin);

            frmIniSettings.Controls.Add(pnl);

            string[] arrSection = this.GetAllSections();

            // Coordinates for new controls. 
            int X = pxMargin;
            int Y = pxMargin;

            // Add controls for section. 
            foreach (string sSection in arrSection)
            {
                Y += pxMargin;

                // Add section label. 
                Label lblSection = new Label();
                lblSection.Text = sSection;
                lblSection.Width = 100;
                lblSection.Location = new Point(X, Y);
                lblSection.Font = new Font(lblSection.Font.FontFamily, lblSection.Font.Size, FontStyle.Bold);
                pnl.Controls.Add(lblSection);

                Y += lblSection.Height;

                string[] arrEntry = this.GetAllEntries(sSection);

                // Add controls for entries. 
                foreach (string sEntry in arrEntry)
                {
                    int iIndexSeparator = sEntry.IndexOf('=');

                    string sIdentifier = "";
                    string sValue = "";

                    if (iIndexSeparator >= 0) // Separator exists. 
                    {
                        sIdentifier = sEntry.Substring(0, iIndexSeparator);
                        sValue = sEntry.Substring(iIndexSeparator + 1, sEntry.Length - (iIndexSeparator + 1));
                    }
                    else // No separator exists. 
                    {
                        sIdentifier = sEntry;
                    }

                    // Add entry label. 
                    Label lblEntry = new Label();
                    lblEntry.Text = sIdentifier;
                    lblEntry.Width = 150;
                    lblEntry.Location = new Point(X, Y);
                    pnl.Controls.Add(lblEntry);

                    // Add entry textbox. 
                    TextBox txtEntry = new TextBox();
                    txtEntry.Text = sValue;
                    txtEntry.Width = pnl.Width - (X + lblEntry.Width + pxMargin);
                    txtEntry.Location = new Point(X + lblEntry.Width + pxMargin, Y);
                    txtEntry.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
                    txtEntry.TextChanged += TxtEntry_TextChanged;
                    txtEntry.Tag = sSection + "=" + sIdentifier;
                    pnl.Controls.Add(txtEntry);

                    Y += txtEntry.Height + pxMargin;

                    if (string.IsNullOrWhiteSpace(sValue)) // No value for entry. 
                        txtEntry.Enabled = false;
                }
            }

            frmIniSettings.ShowDialog();
        }

        #endregion Methods
        /*****************************************************************/
        // Events
        /*****************************************************************/
        #region Events

        private void TxtEntry_TextChanged(object sender, EventArgs e)
        {
            TextBox txtSender = (TextBox)sender;
            string sTag = txtSender.Tag.ToString();

            string sSection = "";
            string sEntry = "";

            int iIndexSeparator = sTag.IndexOf('=');

            sSection = sTag.Substring(0, iIndexSeparator);
            sEntry = sTag.Substring(iIndexSeparator + 1, sTag.Length - (iIndexSeparator + 1));

            this.Set(sSection, sEntry, txtSender.Text);
        }

        private void FrmIniSettings_KeyDown(object sender, KeyEventArgs e)
        {
            Form frmSender = (Form)sender;

            if (e.KeyCode == Keys.Escape)
                frmSender.Close();
        }

        #endregion Events
    }
}
