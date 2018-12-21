using Gurock.TestRail;
using Inflectra.SpiraTest.AddOns.TestRailImporter.SpiraSoapService;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace Inflectra.SpiraTest.AddOns.TestRailImporter
{
    public class ImportThread
    {
        //Project role for new users
        private const int PROJECT_ROLE_ID = 5;

        protected ProgressForm progressForm;
        protected int testRailProjectId;

        protected Dictionary<int, int> testSuiteMapping;
        protected Dictionary<int, int> testSuiteSectionMapping;
        protected Dictionary<int, int> testCaseMapping;
        protected Dictionary<int, int> testStepMapping;
        protected Dictionary<int, int> testSetMapping;
        protected Dictionary<int, int> testSetTestCaseMapping;
        protected Dictionary<int, int> usersMapping;
        protected Dictionary<int, int> testRunStepMapping;
        protected Dictionary<int, int> testRunStepToTestStepMapping;
        protected Dictionary<int, int> releaseMapping = new Dictionary<int, int>();
        protected int newProjectId;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressForm">Handle to the parent form</param>
        public ImportThread(ProgressForm progressForm, int testRailProjectId)
        {
            this.progressForm = progressForm;
            this.testRailProjectId = testRailProjectId;
        }

        /// <summary>
        /// Configure the SOAP connection for HTTP or HTTPS depending on what was specified
        /// </summary>
        /// <param name="httpBinding"></param>
        /// <param name="uri"></param>
        /// <remarks>Allows self-signed certs to be used</remarks>
        public void ConfigureBinding(BasicHttpBinding httpBinding, Uri uri)
        {
            //Handle SSL if necessary
            if (uri.Scheme == "https")
            {
                httpBinding.Security.Mode = BasicHttpSecurityMode.Transport;
                httpBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.None;

                //Allow self-signed certificates
                PermissiveCertificatePolicy.Enact("");
            }
            else
            {
                httpBinding.Security.Mode = BasicHttpSecurityMode.None;
            }
            httpBinding.AllowCookies = true;
        }

        private void ImportProject(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, APIClient testRailApi)
        {
            try
            {
                streamWriter.WriteLine(String.Format("Importing Test Rail project {0} to Spira...", this.testRailProjectId));

                //Get the project info from the TestRail API
                JObject project = (JObject)testRailApi.SendGet("get_project/" + this.testRailProjectId);
                if (project == null)
                {
                    throw new ApplicationException(String.Format("Empty project data returned by TestRail for project '{0}' so aborting import", this.testRailProjectId));
                }

                RemoteProject remoteProject = new RemoteProject();
                remoteProject.Name = project["name"].Value<string>();
                remoteProject.Website = project["url"].Value<string>();
                remoteProject.Active = !project["is_completed"].Value<bool>();

                //Reconnect and import the project
                spiraClient.Connection_Authenticate2(Properties.Settings.Default.SpiraUserName, Properties.Settings.Default.SpiraPassword, "TestRailImporter");
                remoteProject = spiraClient.Project_Create(remoteProject, null);
                streamWriter.WriteLine("New Project '" + remoteProject.Name + "' Created");
                int projectId = remoteProject.ProjectId.Value;

                this.newProjectId = remoteProject.ProjectId.Value;
            }
            catch (APIException exception)
            {
                streamWriter.WriteLine("Unable to access the TestRail API. The error message is: " + exception.Message);
                throw;
            }
            catch (Exception exception)
            {
                streamWriter.WriteLine("General error importing data from TestRail to Spira The error message is: " + exception.Message);
                throw;
            }
        }

        private void ImportUsers(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, APIClient testRailApi, int userId)
        {
            //Get the users from the TestRail API
            JArray users = (JArray)testRailApi.SendGet("get_users");
            if (users != null)
            {
                foreach (JObject user in users)
                {
                    //Extract the user data
                    int testRailId = user["id"].Value<int>();
                    string userName = user["email"].Value<string>();
                    string fullName = user["name"].Value<string>();
                    string firstName = userName;
                    string lastName = userName;
                    if (fullName.Contains(' '))
                    {
                        firstName = fullName.Substring(0, fullName.IndexOf(' '));
                        lastName = fullName.Substring(fullName.IndexOf(' ') + 1);
                    }
                    else
                    {
                        firstName = fullName;
                        lastName = fullName;
                    }
                    string emailAddress = userName;
                    bool isActive = user["is_active"].Value<bool>();

                    //See if we're importing users or mapping to a single user
                    if (Properties.Settings.Default.Users)
                    {
                        //Default to observer role for all imports, for security reasons
                        RemoteUser remoteUser = new RemoteUser();
                        remoteUser.FirstName = firstName;
                        remoteUser.LastName = lastName;
                        remoteUser.UserName = userName;
                        remoteUser.EmailAddress = emailAddress;
                        remoteUser.Active = isActive;
                        remoteUser.Approved = true;
                        remoteUser.Admin = false;
                        userId = spiraClient.User_Create(remoteUser, Properties.Settings.Default.NewUserPassword, "What was my default email address?", emailAddress, PROJECT_ROLE_ID).UserId.Value;
                    }

                    //Add the mapping to the hashtable for use later on
                    this.usersMapping.Add(testRailId, userId);
                }
            }
        }

        private void ImportReleases(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, APIClient testRailApi)
        {
            //Get the milestones from the TestRail API
            JArray milestones = (JArray)testRailApi.SendGet("get_milestones/" + this.testRailProjectId);
            if (milestones != null)
            {
                foreach (JObject milestone in milestones)
                {
                    try
                    {
                        //Extract the user data
                        int testRailId = milestone["id"].Value<int>();

                        //Load the release and capture the ID
                        RemoteRelease remoteRelease = new RemoteRelease();
                        remoteRelease.Name = milestone["name"].Value<string>();
                        remoteRelease.Description = (milestone["description"] == null) ? null : milestone["description"].Value<string>();
                        remoteRelease.VersionNumber = testRailId.ToString();    //Use the test rail ID as the 'version number
                        remoteRelease.StartDate = (milestone["start_on"] == null) ? DateTime.UtcNow : FromUnixTime(milestone["start_on"].Value<long>());
                        remoteRelease.EndDate = (milestone["due_on"] == null) ? DateTime.UtcNow.AddMonths(1) : FromUnixTime(milestone["due_on"].Value<long>());
                        remoteRelease.ResourceCount = 1;
                        remoteRelease.ReleaseTypeId = /* Major Release */1;
                        remoteRelease.ReleaseStatusId = /* Planned */1;
                        bool isStarted = (milestone["is_started"] == null) ? false : milestone["is_started"].Value<bool>();
                        bool isCompleted = (milestone["is_completed"] == null) ? false : milestone["is_completed"].Value<bool>();
                        if (isCompleted)
                        {
                            remoteRelease.ReleaseStatusId = /* Completed */3;
                        }
                        else if (isStarted)
                        {
                            remoteRelease.ReleaseStatusId = /* In Progress */2;
                        }

                        //See if we have a matching parent release
                        int? parentReleaseId = null;
                        if (milestone["parent_id"] != null)
                        {
                            int? testRailParentMilestoneId = milestone["parent_id"].Value<int?>();
                            if (testRailParentMilestoneId.HasValue && this.releaseMapping.ContainsKey(testRailParentMilestoneId.Value))
                            {
                                parentReleaseId = this.releaseMapping[testRailParentMilestoneId.Value];
                            }
                        }

                        int newReleaseId = spiraClient.Release_Create(remoteRelease, parentReleaseId).ReleaseId.Value;
                        streamWriter.WriteLine("Added release: " + testRailId);

                        //Add to the mapping hashtables
                        if (!this.releaseMapping.ContainsKey(testRailId))
                        {
                            this.releaseMapping.Add(testRailId, newReleaseId);
                        }
                    }
                    catch (Exception ex)
                    {
                        streamWriter.WriteLine("Ignoring TestRail milestone: " + milestone.ToString() + " - " + ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Imports the test plans
        /// </summary>
        /// <param name="streamWriter"></param>
        /// <param name="spiraClient"></param>
        /// <param name="testRailApi"></param>
        private void ImportTestPlans(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, APIClient testRailApi)
        {
            //Get the test cases from the TestRail API
            JArray testPlans = (JArray)testRailApi.SendGet("get_plans/" +this.testRailProjectId);
            if (testPlans != null)
            {
                foreach (JObject testPlan in testPlans)
                {
                    //Extract the test rail data
                    int testRailId = testPlan["id"].Value<int>();
                    bool isCompleted = testPlan["is_completed"].Value<bool>();

                    //Create the new SpiraTest test set
                    RemoteTestSet remoteTestSet = new RemoteTestSet();
                    remoteTestSet.Name = testPlan["name"].Value<string>();
                    remoteTestSet.Description = testPlan["description"].Value<string>();
                    remoteTestSet.TestSetStatusId = (isCompleted) ? /* Completed */ 3 : /* Not Started */ 1;
                    remoteTestSet.TestRunTypeId = 1; /* Manual */
                    int? createdById = testPlan["created_by"].Value<int?>();
                    if (createdById.HasValue && this.usersMapping.ContainsKey(createdById.Value))
                    {
                        remoteTestSet.CreatorId = this.usersMapping[createdById.Value];
                    }

                    int newTestSetId = spiraClient.TestSet_Create(remoteTestSet).TestSetId.Value;
                    streamWriter.WriteLine("Added test plan: " + testRailId);

                    //Add to the mapping dictionary
                    if (!this.testSetMapping.ContainsKey(testRailId))
                    {
                        this.testSetMapping.Add(testRailId, newTestSetId);
                    }

                    //Test Rail doesn't map test cases to test plans, so we don't map them
                }
            }
        }

        /// <summary>
        /// Imports the test runs and 'tests'
        /// </summary>
        /// <param name="streamWriter"></param>
        /// <param name="spiraClient"></param>
        /// <param name="testRailApi"></param>
        private void ImportTestRuns(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, APIClient testRailApi)
        {
            //Get the test runs that are not part of a plan
            JArray testRuns = (JArray)testRailApi.SendGet("get_runs/" + this.testRailProjectId);
            if (testRuns != null)
            {
                foreach (JObject testRun in testRuns)
                {
                    ImportTestRun(streamWriter, spiraClient, testRailApi, testRun);
                }
            }

            //Next get the test runs for each plan
            foreach(KeyValuePair<int, int> kvp in this.testSetMapping)
            {
                //Get the test runs that are part of a plan
                int testPlanId = kvp.Key;
                JObject testPlan = (JObject)testRailApi.SendGet("get_plan/" + testPlanId);
                testRuns = (JArray)testPlan["entries"];
                if (testRuns != null)
                {
                    foreach (JObject testRun in testRuns)
                    {
                        ImportTestRun(streamWriter, spiraClient, testRailApi, testRun);
                    }
                }
            }
        }

        /// <summary>
        /// Imports a single test run with its tests
        /// </summary>
        private void ImportTestRun(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, APIClient testRailApi, JObject testRun)
        {
            //Extract the test rail data
            if (testRun["id"].Type == JTokenType.Integer)
            {
                int testRailRunId = testRun["id"].Value<int>();
                streamWriter.WriteLine("Importing test run: " + testRailRunId);

                int? milestoneId = testRun["milestone_id"].Value<int?>();
                int? planId = testRun["plan_id"].Value<int?>();
                string name = testRun["name"].Value<string>();
                string description = testRun["description"].Value<string>();

                //See if we have any tests in the run
                JArray tests = (JArray)testRailApi.SendGet("get_tests/" + testRailRunId);
                if (tests != null)
                {
                    foreach (JObject test in tests)
                    {
                        int trTestId = test["id"].Value<int>();
                        streamWriter.WriteLine("Importing test: T" + trTestId);
                        int trTestCaseId = test["case_id"].Value<int>();
                        int trStatusId = test["status_id"].Value<int>();
                        string title = test["title"].Value<string>();

                        int? spiraTesterId = null;
                        int? spiraReleaseId = null;
                        int spiraTestCaseId = this.testCaseMapping[trTestCaseId];
                        int? spiraTestSetId = null;
                        int spiraExecutionStatusId = ConvertExecutionStatus(trStatusId);

                        int? assignedToId = test["assignedto_id"].Value<int?>();
                        if (assignedToId.HasValue && this.usersMapping.ContainsKey(assignedToId.Value))
                        {
                            spiraTesterId = this.usersMapping[assignedToId.Value];
                        }
                        if (milestoneId.HasValue && this.releaseMapping.ContainsKey(milestoneId.Value))
                        {
                            spiraReleaseId = this.releaseMapping[milestoneId.Value];
                        }
                        if (planId.HasValue && this.testSetMapping.ContainsKey(planId.Value))
                        {
                            spiraTestSetId = this.testSetMapping[planId.Value];
                        }

                        //Now get the results for the test
                        JArray results = (JArray)testRailApi.SendGet("get_results/" + trTestId);
                        if (results != null)
                        {
                            foreach (JObject result in results)
                            {
                                int trResultId = result["id"].Value<int>();
                                streamWriter.WriteLine("Importing result: " + trResultId);

                                DateTime executionDate = FromUnixTime(result["created_on"].Value<long>());
                                //streamWriter.WriteLine("*DEBUG*: " + executionDate.ToString() + "==" + result["created_on"].Value<long>());
                                string actualResults = result["comment"].Value<string>();

                                //We now upload each of these as a test run to SpiraTest

                                //Now create a new test run shell from this test case
                                RemoteManualTestRun[] remoteTestRuns = spiraClient.TestRun_CreateFromTestCases(new int[] { spiraTestCaseId }, null);
                                if (remoteTestRuns.Length > 0)
                                {
                                    //We concatenate the test run name and test title from TestRail
                                    string testRunName = name + ": " + title;

                                    //Update the test run information
                                    RemoteManualTestRun remoteTestRun = remoteTestRuns[0];
                                    remoteTestRun.ExecutionStatusId = spiraExecutionStatusId;
                                    remoteTestRun.StartDate = executionDate;
                                    DateTime endDate = executionDate;   //We don't have a good way to parse TestRail durations
                                    remoteTestRun.EndDate = endDate;
                                    remoteTestRun.Name = testRunName;
                                    remoteTestRun.TestSetId = spiraTestSetId;
                                    //remoteTestRun.TestSetTestCaseId = testSetTestCaseId;  //No equivalent in TestRail
                                    remoteTestRun.TesterId = spiraTesterId;
                                    remoteTestRun.ReleaseId = spiraReleaseId;

                                    //Now the steps
                                    foreach (RemoteTestRunStep remoteTestRunStep in remoteTestRun.TestRunSteps)
                                    {
                                        remoteTestRunStep.ActualResult = actualResults;
                                        remoteTestRunStep.ExecutionStatusId = spiraExecutionStatusId;
                                    }

                                    //Save the run
                                    spiraClient.TestRun_Save(remoteTestRuns, executionDate);
                                    streamWriter.WriteLine("Added manual test run for result: " + trResultId);
                                }
                            }
                        }
                    }
                }
            }
        }

        public static DateTime FromUnixTime(long unixTime)
        {
            return epoch.AddSeconds(unixTime);
        }
        private static readonly DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        /// <summary>
        /// Converts Test Rail execution statuses
        /// </summary>
        private int ConvertExecutionStatus(int testRailStatusId)
        {
            int executionStatusId;
            switch (testRailStatusId)
            {
                case 1:
                    executionStatusId = (int)ExecutionStatusEnum.Passed;
                    break;
                case 2:
                    executionStatusId = (int)ExecutionStatusEnum.Blocked;
                    break;
                case 3:
                    /* Untested*/
                    executionStatusId = (int)ExecutionStatusEnum.NotRun;
                    break;
                case 4:
                    /* Retest*/
                    executionStatusId = (int)ExecutionStatusEnum.Caution;
                    break;
                case 5:
                    /* Failed*/
                    executionStatusId = (int)ExecutionStatusEnum.Failed;
                    break;

                default:
                    //Custom statuses become N/A
                    executionStatusId = (int)ExecutionStatusEnum.NotApplicable;
                    break;

            }

            return executionStatusId;
        }

        /// <summary>
        /// The execution statuses
        /// </summary>
        public enum ExecutionStatusEnum
        {
            Failed = 1,
            Passed = 2,
            NotRun = 3,
            NotApplicable = 4,
            Blocked = 5,
            Caution = 6
        }

        /// <summary>
        /// Imports the test suites and sections
        /// </summary>
        /// <param name="streamWriter"></param>
        /// <param name="spiraClient"></param>
        /// <param name="testRailApi"></param>
        private void ImportTestSuites(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, APIClient testRailApi)
        {
            //Get the test cases from the TestRail API
            JArray testSuites = (JArray)testRailApi.SendGet("get_suites/" + this.testRailProjectId);
            if (testSuites != null)
            {
                foreach (JObject testSuite in testSuites)
                {
                    //Extract the user data
                    int testRailId = testSuite["id"].Value<int>();

                    //Create the new SpiraTest test folder
                    RemoteTestCaseFolder remoteTestFolder = new RemoteTestCaseFolder();
                    remoteTestFolder.Name = testSuite["name"].Value<string>();
                    remoteTestFolder.Description = testSuite["description"].Value<string>();
                    int newTestCaseFolderId = spiraClient.TestCase_CreateFolder(remoteTestFolder).TestCaseFolderId.Value;
                    streamWriter.WriteLine("Added test suite: " + testRailId);

                    //Add to the mapping hashtables
                    if (!this.testSuiteMapping.ContainsKey(testRailId))
                    {
                        this.testSuiteMapping.Add(testRailId, newTestCaseFolderId);
                    }

                    //Now get each section in the suite
                    JArray testSections = (JArray)testRailApi.SendGet("get_sections/" +this.testRailProjectId + "&suite_id=" + testRailId);
                    if (testSections != null)
                    {
                        Dictionary<string, int> usedSectionNames = new  Dictionary<string, int>();
                        foreach (JObject testSection in testSections)
                        {
                            //Extract the user data
                            int testRailSectionId = testSection["id"].Value<int>();

                            //See if we have a name for this already in this suite, if so, re-use the existing section name
                            string sectionName = testSection["name"].Value<string>();
                            if (usedSectionNames.ContainsKey(sectionName))
                            {
                                //Add to the mapping dictionary
                                if (!this.testSuiteSectionMapping.ContainsKey(testRailSectionId))
                                {
                                    this.testSuiteSectionMapping.Add(testRailSectionId, usedSectionNames[sectionName]);
                                }
                            }
                            else
                            {
                                //Create the new SpiraTest test folder
                                remoteTestFolder = new RemoteTestCaseFolder();
                                remoteTestFolder.Name = sectionName;
                                remoteTestFolder.Description = testSection["description"].Value<string>();

                                //See if we have a parent section already imported
                                int? parentSectionId = testSection["parent_id"].Value<int?>();
                                if (parentSectionId.HasValue && this.testSuiteSectionMapping.ContainsKey(parentSectionId.Value))
                                {
                                    remoteTestFolder.ParentTestCaseFolderId = this.testSuiteSectionMapping[parentSectionId.Value];
                                }
                                else
                                {
                                    //Otherwise just import directly under the suite
                                    remoteTestFolder.ParentTestCaseFolderId = newTestCaseFolderId;
                                }

                                int newTestCaseFolderId2 = spiraClient.TestCase_CreateFolder(remoteTestFolder).TestCaseFolderId.Value;
                                streamWriter.WriteLine("Added test section: " + testRailSectionId);
                                usedSectionNames.Add(sectionName, newTestCaseFolderId2);

                                //Add to the mapping dictionary
                                if (!this.testSuiteSectionMapping.ContainsKey(testRailSectionId))
                                {
                                    this.testSuiteSectionMapping.Add(testRailSectionId, newTestCaseFolderId2);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Imports test cases and associated steps
        /// </summary>
        /// <param name="streamWriter"></param>
        /// <param name="spiraClient"></param>
        /// <param name="testRailApi"></param>
        private void ImportTestCases(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, APIClient testRailApi)
        {
            //If we have any test suites, then we need to query by suite, otherwise we can get all for the project
            foreach (KeyValuePair<int, int> kvp in this.testSuiteMapping)
            {
                //Get the test cases from the TestRail API
                JArray testCases = (JArray)testRailApi.SendGet("get_cases/" +this.testRailProjectId + "&suite_id=" + kvp.Key);
                if (testCases != null)
                {
                    foreach (JObject testCase in testCases)
                    {
                        //Extract the user data
                        int testRailId = testCase["id"].Value<int>();

                        //Create the new SpiraTest test case
                        RemoteTestCase remoteTestCase = new RemoteTestCase();
                        remoteTestCase.Name = testCase["title"].Value<string>();
                        if (testCase["custom_description"] != null && testCase["custom_preconds"] != null)
                        {
                            string description = testCase["custom_description"].Value<string>();
                            string preconds = testCase["custom_preconds"].Value<string>();
                            remoteTestCase.Description = description + " " + preconds;
                        }
                        remoteTestCase.TestCaseTypeId = /* Functional */ 3;
                        remoteTestCase.TestCaseStatusId = /* Ready For Test */ 5;
                        if (testCase["priority_id"] != null)
                        {
                            int? testCasePriorityId = testCase["priority_id"].Value<int?>();
                            if (testCasePriorityId.HasValue && testCasePriorityId >= 1 && testCasePriorityId <= 4)
                            {
                                remoteTestCase.TestCasePriorityId = testCasePriorityId.Value;
                            }
                        }
                        if (testCase["created_by"] != null)
                        {
                            int? createdById = testCase["created_by"].Value<int?>();
                            if (createdById.HasValue && this.usersMapping.ContainsKey(createdById.Value))
                            {
                                remoteTestCase.AuthorId = this.usersMapping[createdById.Value];
                            }
                        }

                        //See if we have a suite or section to put it under (section takes precedence)
                        int? testSuiteId = null;
                        if (testCase["suite_id"] != null)
                        {
                            testSuiteId = testCase["suite_id"].Value<int?>();
                        }
                        int? testSectionId = null;
                        if (testCase["section_id"] != null)
                        {
                            testSectionId = testCase["section_id"].Value<int?>();
                        }

                        if (testSectionId.HasValue && this.testSuiteSectionMapping.ContainsKey(testSectionId.Value))
                        {
                            remoteTestCase.TestCaseFolderId = this.testSuiteSectionMapping[testSectionId.Value];
                        }
                        else if (testSuiteId.HasValue && this.testSuiteMapping.ContainsKey(testSuiteId.Value))
                        {
                            remoteTestCase.TestCaseFolderId = this.testSuiteMapping[testSuiteId.Value];
                        }
                        int newTestCaseId = spiraClient.TestCase_Create(remoteTestCase).TestCaseId.Value;
                        streamWriter.WriteLine("Added test case: " + testRailId);

                        //Add the test case to the release (if specified)
                        int? milestoneId = testCase["milestone_id"].Value<int?>();
                        if (milestoneId.HasValue && this.releaseMapping.ContainsKey(milestoneId.Value))
                        {
                            int releaseId = this.releaseMapping[milestoneId.Value];
                            spiraClient.Release_AddTestMapping(new RemoteReleaseTestCaseMapping() { ReleaseId = releaseId, TestCaseId = newTestCaseId });
                        }

                        //Now see if we have steps (custom steps in TestRail)
                        JArray testSteps = testCase["custom_steps_separated"].Value<JArray>();
                        if (testSteps != null && testSteps.Count > 0)
                        {
                            streamWriter.WriteLine("Adding " + testSteps.Count + " custom steps to test case: " + testRailId);
                            foreach (JObject testStep in testSteps)
                            {
                                RemoteTestStep remoteTestStep = new RemoteTestStep();
                                remoteTestStep.Description = testStep["content"].Value<string>();
                                remoteTestStep.ExpectedResult = testStep["expected"].Value<string>();
                                spiraClient.TestCase_AddStep(remoteTestStep, newTestCaseId);
                            }
                        }
                        streamWriter.WriteLine("Added custom steps to test case: " + testRailId);

                        //Add to the mapping hashtables
                        if (!this.testCaseMapping.ContainsKey(testRailId))
                        {
                            this.testCaseMapping.Add(testRailId, newTestCaseId);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This method is responsible for actually importing the data
        /// </summary>
        /// <param name="stateInfo">State information handle</param>
        /// <remarks>This runs in background thread to avoid freezing the progress form</remarks>
        public void ImportData(object stateInfo)
        {
            //First open up the textfile that we will log information to (used for debugging purposes)
            string debugFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Spira_TestRail_Import.log";
            StreamWriter streamWriter = File.CreateText(debugFile);

            try
            {
                streamWriter.WriteLine("Starting import at: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());

                //Create the mapping hashtables
                this.testCaseMapping = new Dictionary<int, int>();
                this.testSuiteMapping = new Dictionary<int, int>();
                this.testSuiteSectionMapping = new Dictionary<int, int>();
                this.testStepMapping = new Dictionary<int, int>();
                this.usersMapping = new Dictionary<int, int>();
                this.testRunStepMapping = new Dictionary<int, int>();
                this.testRunStepToTestStepMapping = new Dictionary<int, int>();
                this.testSetMapping = new Dictionary<int, int>();
                this.testSetTestCaseMapping = new Dictionary<int, int>();

                //Connect to Spira
                streamWriter.WriteLine("Connecting to Spira...");
                SpiraSoapService.SoapServiceClient spiraClient = new SpiraSoapService.SoapServiceClient();

                //Set the end-point and allow cookies
                string url = Properties.Settings.Default.SpiraUrl + ImportForm.IMPORT_WEB_SERVICES_URL;
                spiraClient.Endpoint.Address = new EndpointAddress(url);
                BasicHttpBinding httpBinding = (BasicHttpBinding)spiraClient.Endpoint.Binding;
                ConfigureBinding(httpBinding, spiraClient.Endpoint.Address.Uri);

                bool success = spiraClient.Connection_Authenticate2(Properties.Settings.Default.SpiraUserName, Properties.Settings.Default.SpiraPassword, "TestRailImporter");
                if (!success)
                {
                    string message = "Failed to authenticate with Spira using login: '" + Properties.Settings.Default.SpiraUserName + "' so terminating import!";
                    streamWriter.WriteLine(message);
                    streamWriter.Close();

                    //Display the exception message
                    this.progressForm.ProgressForm_OnError(message);
                    return;
                }

                //Connect to TestRail
                APIClient testRailApi = new APIClient(Properties.Settings.Default.TestRailUrl);
                testRailApi.User = Properties.Settings.Default.TestRailUserName;
                testRailApi.Password = Properties.Settings.Default.TestRailPassword;

                //1) Create a new project
                ImportProject(streamWriter, spiraClient, testRailApi);

                //2) Get the users and import - if we don't want to import user, map all TestRail users to single SpiraId
                int userId = -1;
                string newPassword = Properties.Settings.Default.NewUserPassword;
                if (!Properties.Settings.Default.Users)
                {
                    RemoteUser remoteUser = new RemoteUser();
                    remoteUser.FirstName = "Test";
                    remoteUser.LastName = "Rail";
                    remoteUser.UserName = "testrail";
                    remoteUser.EmailAddress = "testrail@mycompany.com";
                    remoteUser.Active = true;
                    remoteUser.Approved = true;
                    remoteUser.Admin = false;
                    userId = spiraClient.User_Create(remoteUser, newPassword, "What was my default email address?", remoteUser.EmailAddress, PROJECT_ROLE_ID).UserId.Value;
                }

                ImportUsers(streamWriter, spiraClient, testRailApi, userId);

                //**** Show that we've imported users ****
                if (Properties.Settings.Default.Users)
                {
                    streamWriter.WriteLine("Users Imported");
                    this.progressForm.ProgressForm_OnProgressUpdate(1);
                }

                //3) Get the releases and import
                if (Properties.Settings.Default.Releases)
                {
                    ImportReleases(streamWriter, spiraClient, testRailApi);
                }

                //**** Show that we've imported releases ****
                streamWriter.WriteLine("Releases Imported");
                this.progressForm.ProgressForm_OnProgressUpdate(2);

                //3-4) Import the test suites and test cases
                if (Properties.Settings.Default.TestCases)
                {
                    ImportTestSuites(streamWriter, spiraClient, testRailApi);

                    //**** Show that we've imported test cases and test steps ****
                    streamWriter.WriteLine("Test Suites Imported");
                    this.progressForm.ProgressForm_OnProgressUpdate(3);

                    ImportTestCases(streamWriter, spiraClient, testRailApi);

                    //**** Show that we've imported test cases and test steps ****
                    streamWriter.WriteLine("Test Cases and Custom Steps Imported");
                    this.progressForm.ProgressForm_OnProgressUpdate(4);
                }

                //5) Import the test plans (check for case import)
                if (Properties.Settings.Default.TestCases)
                {
                    ImportTestPlans(streamWriter, spiraClient, testRailApi);

                    //**** Show that we've imported test cases and test steps ****
                    streamWriter.WriteLine("Test Plans Imported");
                    this.progressForm.ProgressForm_OnProgressUpdate(5);
                }

                //6) Import the test runs (have to have case as well!)
                if (Properties.Settings.Default.TestCases && Properties.Settings.Default.TestRuns)
                {
                    ImportTestRuns(streamWriter, spiraClient, testRailApi);

                    //**** Show that we've imported test cases and test steps ****
                    streamWriter.WriteLine("Test Runs Imported");
                    this.progressForm.ProgressForm_OnProgressUpdate(6);
                }

                //**** Mark the form as finished ****
                streamWriter.WriteLine("Import completed at: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
                this.progressForm.ProgressForm_OnFinish();

                //Close the debugging file
                streamWriter.Close();
            }
            catch (Exception exception)
            {
                //Log the error
                streamWriter.WriteLine("*ERROR* Occurred during Import: '" + exception.Message + "' at " + exception.Source + " (" + exception.StackTrace + ")");
                streamWriter.Close();

                //Display the exception message
                this.progressForm.ProgressForm_OnError(exception);
            }
        }
    }
}
