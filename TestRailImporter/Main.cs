using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Linq;

using Gurock.TestRail;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace Inflectra.SpiraTest.AddOns.TestRailImporter
{
	/// <summary>
	/// This is the code behind class for the utility that imports projects from
	/// HP Mercury Quality Center / TestDirector into Inflectra SpiraTest
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboProject;
		private System.Windows.Forms.TextBox txtPassword;
		private System.Windows.Forms.TextBox txtLogin;
        private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnAuthenticate;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtServer;
		private System.Windows.Forms.Button btnNext;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		protected ImportForm importForm;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		public System.Windows.Forms.CheckBox chkImportRequirements;
		public System.Windows.Forms.CheckBox chkImportTestCases;
        public System.Windows.Forms.CheckBox chkImportTestRuns;
        public System.Windows.Forms.CheckBox chkImportUsers;
        private CheckBox chkPassword;
		protected ProgressForm progressForm;

        /// <summary>
        /// The id of the selected test rail project
        /// </summary>
        public int TestRailProjectId
        {
            get;
            set;
        }

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			// Add any event handlers
			this.Closing += new CancelEventHandler(MainForm_Closing);

			//Set the initial state of any buttons
			this.btnNext.Enabled = false;

			//Create the other forms and set a handle to this form and the import form
			this.importForm = new ImportForm();
			this.progressForm = new ProgressForm();
			this.importForm.MainFormHandle = this;
			this.importForm.ProgressFormHandle = this.progressForm;
			this.progressForm.MainFormHandle = this;
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.btnNext = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkPassword = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnAuthenticate = new System.Windows.Forms.Button();
            this.cboProject = new System.Windows.Forms.ComboBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtLogin = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkImportUsers = new System.Windows.Forms.CheckBox();
            this.chkImportTestRuns = new System.Windows.Forms.CheckBox();
            this.chkImportTestCases = new System.Windows.Forms.CheckBox();
            this.chkImportRequirements = new System.Windows.Forms.CheckBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(451, 74);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(115, 26);
            this.btnNext.TabIndex = 0;
            this.btnNext.Text = "Next >";
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(326, 74);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(106, 26);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(19, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(528, 27);
            this.label1.TabIndex = 6;
            this.label1.Text = "1. Connect to TestRail";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkPassword);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtServer);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnAuthenticate);
            this.groupBox1.Controls.Add(this.cboProject);
            this.groupBox1.Controls.Add(this.txtPassword);
            this.groupBox1.Controls.Add(this.txtLogin);
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(29, 55);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(576, 240);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "TestRail Configuration";
            // 
            // chkPassword
            // 
            this.chkPassword.AutoSize = true;
            this.chkPassword.Location = new System.Drawing.Point(115, 128);
            this.chkPassword.Name = "chkPassword";
            this.chkPassword.Size = new System.Drawing.Size(164, 21);
            this.chkPassword.TabIndex = 22;
            this.chkPassword.Text = "Remember Password";
            this.chkPassword.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(29, 28);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 18);
            this.label6.TabIndex = 21;
            this.label6.Text = "Server:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtServer
            // 
            this.txtServer.Location = new System.Drawing.Point(115, 28);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(403, 22);
            this.txtServer.TabIndex = 20;
            this.txtServer.Text = "http://myserver/qcbin";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(19, 175);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(154, 19);
            this.label5.TabIndex = 19;
            this.label5.Text = "Project:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(28, 103);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 18);
            this.label3.TabIndex = 17;
            this.label3.Text = "Password:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(10, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(96, 18);
            this.label2.TabIndex = 16;
            this.label2.Text = "User Name:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // btnAuthenticate
            // 
            this.btnAuthenticate.Location = new System.Drawing.Point(413, 128);
            this.btnAuthenticate.Name = "btnAuthenticate";
            this.btnAuthenticate.Size = new System.Drawing.Size(105, 27);
            this.btnAuthenticate.TabIndex = 14;
            this.btnAuthenticate.Text = "Authenticate";
            this.btnAuthenticate.Click += new System.EventHandler(this.btnAuthenticate_Click);
            // 
            // cboProject
            // 
            this.cboProject.Location = new System.Drawing.Point(182, 175);
            this.cboProject.Name = "cboProject";
            this.cboProject.Size = new System.Drawing.Size(336, 24);
            this.cboProject.TabIndex = 13;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(115, 100);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(402, 22);
            this.txtPassword.TabIndex = 11;
            // 
            // txtLogin
            // 
            this.txtLogin.Location = new System.Drawing.Point(115, 65);
            this.txtLogin.Name = "txtLogin";
            this.txtLogin.Size = new System.Drawing.Size(403, 22);
            this.txtLogin.TabIndex = 10;
            this.txtLogin.Text = "alex_alm";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkImportUsers);
            this.groupBox2.Controls.Add(this.btnCancel);
            this.groupBox2.Controls.Add(this.chkImportTestRuns);
            this.groupBox2.Controls.Add(this.chkImportTestCases);
            this.groupBox2.Controls.Add(this.chkImportRequirements);
            this.groupBox2.Controls.Add(this.btnNext);
            this.groupBox2.ForeColor = System.Drawing.Color.Black;
            this.groupBox2.Location = new System.Drawing.Point(29, 314);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(576, 111);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Import Options";
            // 
            // chkImportUsers
            // 
            this.chkImportUsers.Checked = true;
            this.chkImportUsers.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportUsers.Location = new System.Drawing.Point(202, 55);
            this.chkImportUsers.Name = "chkImportUsers";
            this.chkImportUsers.Size = new System.Drawing.Size(105, 28);
            this.chkImportUsers.TabIndex = 6;
            this.chkImportUsers.Text = "Users";
            // 
            // chkImportTestRuns
            // 
            this.chkImportTestRuns.Location = new System.Drawing.Point(202, 28);
            this.chkImportTestRuns.Name = "chkImportTestRuns";
            this.chkImportTestRuns.Size = new System.Drawing.Size(220, 27);
            this.chkImportTestRuns.TabIndex = 4;
            this.chkImportTestRuns.Text = "Test Runs";
            // 
            // chkImportTestCases
            // 
            this.chkImportTestCases.Checked = true;
            this.chkImportTestCases.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportTestCases.Location = new System.Drawing.Point(19, 55);
            this.chkImportTestCases.Name = "chkImportTestCases";
            this.chkImportTestCases.Size = new System.Drawing.Size(221, 28);
            this.chkImportTestCases.TabIndex = 3;
            this.chkImportTestCases.Text = "Test Cases";
            this.chkImportTestCases.CheckedChanged += new System.EventHandler(this.chkImportTestCases_CheckedChanged);
            // 
            // chkImportRequirements
            // 
            this.chkImportRequirements.Checked = true;
            this.chkImportRequirements.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportRequirements.Location = new System.Drawing.Point(19, 28);
            this.chkImportRequirements.Name = "chkImportRequirements";
            this.chkImportRequirements.Size = new System.Drawing.Size(221, 27);
            this.chkImportRequirements.TabIndex = 2;
            this.chkImportRequirements.Text = "Milestones";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Inflectra.SpiraTest.AddOns.TestRailImporter.Properties.Resources.TestRail_Icon;
            this.pictureBox1.Location = new System.Drawing.Point(566, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(48, 46);
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // MainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.ClientSize = new System.Drawing.Size(619, 437);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "SpiraTest Importer for TestRail";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MiniDump.CreateMiniDump();
        }

		/// <summary>
		/// Authenticates the user from the providing server/login/password information
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void btnAuthenticate_Click(object sender, System.EventArgs e)
		{
			//Disable the next button
			this.btnNext.Enabled = false;

			//Make sure that a login was entered
			if (this.txtLogin.Text.Trim() == "")
			{
				MessageBox.Show ("You need to enter a login", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			//Make sure that a server was entered
			if (this.txtServer.Text.Trim() == "")
			{
				MessageBox.Show ("You need to enter the URL to your TestRail instance", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

            try
            {
                //Instantiate the connection to TestRail (by getting a list of projects)
                APIClient testRailApi = new APIClient(this.txtServer.Text.Trim());
                testRailApi.User = this.txtLogin.Text.Trim();
                testRailApi.Password = this.txtPassword.Text.Trim();
                JArray projects = (JArray)testRailApi.SendGet("get_projects");

                if (projects.Count > 0)
                {
                    MessageBox.Show("You have logged into TestRail Successfully", "Authentication", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("You have logged into TestRail successfully, but you don't have any projects to import!", "No Projects Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //Now we need to populate the list of projects
                List<NameIdObject> projectsList = new List<NameIdObject>();
                //Convert into a standard bindable list source
                for (int i = 0; i < projects.Count; i++)
                {
                    NameIdObject project = new NameIdObject();
                    project.Id = projects[i]["id"].Value<string>();
                    project.Name = projects[i]["name"].Value<string>();
                    projectsList.Add(project);
                }

                //Sort by name
                projectsList = projectsList.OrderBy(p => p.Name).ToList();

                this.cboProject.DisplayMember = "Name";
                this.cboProject.DataSource = projectsList;

                //Enable the Next button
                this.btnNext.Enabled = true;
            }
            catch (APIException exception)
            {
                MessageBox.Show("Unable to access the TestRail API. The error message is: " + exception.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            catch (Exception exception)
            {
                MessageBox.Show("General error accessing the TestRail API. The error message is: " + exception.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

		/// <summary>
		/// Closes the application
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			//Close the application
			this.Close();
		}

		/// <summary>
		/// Called if the form is closed
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void MainForm_Closing(object sender, CancelEventArgs e)
		{
            //Nothing needs to be done
		}

		/// <summary>
		/// Called when the Next button is clicked. Switches to the second form
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void btnNext_Click(object sender, System.EventArgs e)
		{
            //Store the info in settings for later
            Properties.Settings.Default.TestRailUrl = this.txtServer.Text.Trim();
            Properties.Settings.Default.TestRailUserName = this.txtLogin.Text.Trim();
            if (chkPassword.Checked)
            {
                Properties.Settings.Default.TestRailPassword = this.txtPassword.Text;  //Don't trip in case it contains a space
            }
            else
            {
                Properties.Settings.Default.TestRailPassword = "";
            }
            Properties.Settings.Default.Releases = this.chkImportRequirements.Checked;
            Properties.Settings.Default.TestCases = this.chkImportTestCases.Checked;
            Properties.Settings.Default.TestRuns = this.chkImportTestRuns.Checked;
            //Properties.Settings.Default.Attachments = this.chkImportAttachments.Checked;
            Properties.Settings.Default.Users = this.chkImportUsers.Checked;
            Properties.Settings.Default.Save();

            //Always put the current password in settings after save, for use in current run
            Properties.Settings.Default.TestRailPassword = this.txtPassword.Text;  //Don't trip in case it contains a space

            //Store the current test rail project id for use by the import thread
            this.TestRailProjectId = Int32.Parse(((NameIdObject)cboProject.SelectedItem).Id);

			//Hide the current form
			this.Hide();

			//Show the second page in the import wizard
			this.importForm.Show();
		}

		/// <summary>
		/// Change the active status of the test run import checkbox depending on this selection
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void chkImportTestCases_CheckedChanged(object sender, System.EventArgs e)
		{
			this.chkImportTestRuns.Enabled = this.chkImportTestCases.Checked;
			this.chkImportTestRuns.Checked = false;
        }

        /// <summary>
        /// Populates the fields when the form is loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            this.txtServer.Text = Properties.Settings.Default.TestRailUrl;
            this.txtLogin.Text = Properties.Settings.Default.TestRailUserName;
            if (String.IsNullOrEmpty(Properties.Settings.Default.TestRailPassword))
            {
                this.chkPassword.Checked = false;
                this.txtPassword.Text = "";
            }
            else
            {
                this.chkPassword.Checked = true;
                this.txtPassword.Text = Properties.Settings.Default.TestRailPassword;
            }

            this.chkImportRequirements.Checked = Properties.Settings.Default.Releases;
            this.chkImportTestCases.Checked = Properties.Settings.Default.TestCases;
            this.chkImportTestRuns.Checked = Properties.Settings.Default.TestRuns;
            //this.chkImportAttachments.Checked = Properties.Settings.Default.Attachments;
            this.chkImportUsers.Checked = Properties.Settings.Default.Users;
        }
	}
}
