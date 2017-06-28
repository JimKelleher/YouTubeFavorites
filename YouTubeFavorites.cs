
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.YouTube.v3.Data;
using Google.Apis.YouTube.v3;

using System.Collections;

using System.Data.SqlClient;
using System.Data;

using System.Diagnostics;
using System.IO;
using System.Reflection;

using System.Threading.Tasks;
using System.Threading;

using System.Windows.Forms;
using System;

namespace YouTubeFavorites
{
    public class YouTubeFavorites
    {
        // BEGIN DATA
        //-----------------------------------------------------------------------------    

        // Google/YouTube API 3 services

        // This variable is filled in two locations, INIT_STEP_1_SignInAndGetYouTubeFavFeed()
        // for read-only access and GetYoutubeServiceDeletionAuth() for deletion access:
        YouTubeService youTubeService;

        // App.Config-supplied stuff:
        public string YOUTUBE_API_BROWSER_KEY;

        public string YOUTUBE_ACCOUNT_CHANNEL;
        public string DATABASE_CONNECTION_STRING;

        public static string RUN_URL_PREFIX, BROWSER;

        // Runtime parameters, 1 of 2.  This is one of the runtime parameters whose
        // values need to be supplied by either a user or a command line argument:
        string strImportFile = ""; // init

        // Sort-order arrays.  These ArrayLists contain all we need to know about a video:
        // Title, VideoId and Duration.

        // NOTE: These "original order" arrays are the interface that holds the data.  They
        // are filled by all of the three data sources (YouTube Favorites, the JimRadio
        // SQL database and file importing):
        private static ArrayList strInDateSavedOrigOrderTitle, strInDateSavedOrigOrderVideoId, strInDateSavedOrigOrderDuration, strInDateSavedOrigOrderItemId;

        // Other sorts:
        private static ArrayList strInTitleOrderTitle,  strInTitleOrderVideoId,  strInTitleOrderDuration,  strInTitleOrderItemId;
        private static ArrayList strInRandomOrderTitle, strInRandomOrderVideoId, strInRandomOrderDuration, strInRandomOrderItemId;

        // Enumerators:
        private enum EnumSortOrder
        {
            DateSaved,
            Title,
            Random
        }

        private enum EnumDataSource
        {
            YouTubeFavorites,
            JimRadio,
            ImportFile
        }

        private enum EnumFormMode
        {
            step1UninitializedForm,
            step2InitializedForm,
            step3SignedInWithList,
            step4AllowDelete,
            step4PreventDelete
        }

        // "State management" of the GUI:
        private static EnumSortOrder enumSortOrderSelected;
        private static EnumDataSource enumDataSourceSelected;

        private static bool boolToggleSelectAll = false;
        private static bool boolRunAutomatically;

        private static int intRowsRetrieved, intRowsDesired;

        // Google/YouTube Atom feed:
        string strFavoritesListId;

        ChannelListResponse channelListResponse;

        // GUI object references:
        private static Form formRef;
        private static DataGridView dataGridViewRef;

        private static GroupBox groupBoxDataSrcRef;
        private static RadioButton radioButtonDataSrcYouTubeFavRef, radioButtonDataSrcJimRadioRef;
        private static RadioButton radioButtonDataSrcImportFileRef;

        private static GroupBox groupBoxSortRef;
        private static RadioButton radioButtonSortDateSavedRef, radioButtonSortTitleRef;
        private static RadioButton radioButtonSortRandomRef;

        private static TextBox textBoxYoutubeFavoritesChannelRef, textBoxSearchRef;
        private static Button buttonSignInAndRetrieveRef, buttonSelectAllRef;
        private static Button buttonPlayRef, buttonDeleteRef;

        private static Button buttonExportFileRef;

        // END DATA
        //-----------------------------------------------------------------------------
        // BEGIN METHODS    

        private string AlterVideoTitle(string strTitle)
        {
            //- - - - - - - - - - - - - - - - - - - - - - - - - - - - 
            // Remove all quotes and double quotes:
            strTitle = strTitle.Replace("'", "");
            strTitle = strTitle.Replace("\"", "");

            // Remove commas because they screw up the CSV export:
            strTitle = strTitle.Replace(",", "");

            //- - - - - - - - - - - - - - - - - - - - - - - - - - - - 
            // NOTE: Videos that have been taken down are showing up
            // with empty Titles:
            if (strTitle.Trim() == "")
            {
                strTitle = "UNKNOWN VIDEO";
            }
            else
            {
               // If the Title starts with "The ":
               if (strTitle.ToLower().Substring(0, 4) == "the ")
                {
                    // Remove it:
                    strTitle = strTitle.Substring(4, strTitle.Length - 4);
                };
            };

            //- - - - - - - - - - - - - - - - - - - - - - - - - - - - 
            // Return the altered (cleaned up) Title:
            return strTitle;
            //- - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        } // core

        public void ExportAsCsvFile()
        {
            string strFileName = string.Empty;

            // Init the get-filename dialog:
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save file as...";
            saveFileDialog1.Filter = "Comma-separated Text files (*.csv)|*.csv|All files (*.*)|*.*";
            saveFileDialog1.RestoreDirectory = true;

            // Get the file name for the export file:
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Write the contents of the current result set to a comma-separated text file,
                // using the StreamWriter.

                // Open the file:
                StreamWriter streamWriter1 = new StreamWriter(saveFileDialog1.FileName);

                int intRowCount = dataGridViewRef.Rows.Count;
                for (int i = 0; i < intRowCount; i++)
                {
                    // Write a row to the file:
                    streamWriter1.WriteLine(dataGridViewRef.Rows[i].Cells[0].Value.ToString() + "," + // title
                                            dataGridViewRef.Rows[i].Cells[1].Value.ToString() + "," + // VideoId
                                            dataGridViewRef.Rows[i].Cells[2].Value.ToString()         // duration
                                           );
                }

                // Close the file:
                streamWriter1.Close();
            }

        } // implementation: CSV text file

        public void FormLoadInit() // PUBLIC
        {
            // This method is called from the form's Form1_Load() event.

            // Get the application-run and video-play URL prefixes and set
            // them as "global" variables:
            Get_App_Config_Constants();

            // This is a one column grid, so no titles are necessary:
            dataGridViewRef.ColumnHeadersVisible = false;
            dataGridViewRef.RowHeadersVisible = false;

            // Multi-row selection is allowed (and encouraged!), nothing else:
            dataGridViewRef.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewRef.AllowUserToAddRows = false;
            dataGridViewRef.AllowUserToDeleteRows = false;
            dataGridViewRef.ReadOnly = true;

            // This column is  wider than the default:
            const int int_TITLE_WIDTH = 320;

            int intColumn;

            // Add columns to the GridView:
            dataGridViewRef.Columns.Add("Title", "");

            intColumn = dataGridViewRef.Columns.Add("Video Id", "");
            dataGridViewRef.Columns[intColumn].Visible = false;

            intColumn = dataGridViewRef.Columns.Add("Seconds", "");
            dataGridViewRef.Columns[intColumn].Visible = false;

            // Init:

            // Size the Title appropriately:
            DataGridViewColumn DataGridViewColumn1 = dataGridViewRef.Columns[0];
            DataGridViewColumn1.Width = int_TITLE_WIDTH;

            // Runtime parameters, 2 of 2.  These are most of the runtime parameters whose
            // values need to be supplied by either a user or a command line argument:
            string strChannel, strSort, strDataSource, strSearch;

            // NOTE: This is a YouTube account's Favorites group of videos.  I use it
            // as a "staging" area for the JimRadio SQL database:
            string strDEMO_CHANNEL   = Properties.Settings.Default.YOUTUBE_ACCOUNT_CHANNEL;

            //bool boolDummyTestCommandLine = false;

            // Init/set defaults:
            strChannel  = strDEMO_CHANNEL;

            strSort = "date";
            strDataSource = "youtube";
            strSearch = "";

            //------------------------------------------------------
            // PART 1, Standard initialization:

            // Set the window title:
            SetFormTitle("YouTube Favorites Player");

            // To keep the GUI implementation simple, we will set these here:
            formRef.AcceptButton = buttonSignInAndRetrieveRef;

            // We are now on this step and in this mode:
            SetFormControlsModally(EnumFormMode.step2InitializedForm);

            //---------------------------------------------------------------------------------------------
            // PART 2, Command-line-driven initialization:
            //- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            // NOTE: To access command line arguements outside of the main() method,
            // use the Environment class:
            string[] strCommandLineArguments = Environment.GetCommandLineArgs();

            // If the program ran from the command line:
            if (strCommandLineArguments.Length > 1)
            //- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

            // Test command lines:

            // Production:
            // "C:\jim kelleher\YouTubeFavoritesPlayer.exe" youtube random JimKelleher "baby, beautiful, love"
            // "C:\jim kelleher\YouTubeFavoritesPlayer.exe" jimradio random "baby, beautiful, love"
            // "C:\jim kelleher\YouTubeFavoritesPlayer.exe" import random YouTubeFavorites.csv "baby, beautiful, love"

            //if (boolDummyTestCommandLine)
            {
                // There is no user:
                boolRunAutomatically = true;

                // Get the argument values:
                strDataSource = strCommandLineArguments[1];
                strSort = strCommandLineArguments[2];

                if (strDataSource == "youtube")
                {
                    strChannel  = strCommandLineArguments[3];

                    try { strSearch = strCommandLineArguments[4]; }
                    catch (Exception) { }
                }

                else if (strDataSource == "jimradio")
                {
                    try { strSearch = strCommandLineArguments[3]; }
                    catch (Exception) { }
                }

                else if (strDataSource == "import")
                {
                    strImportFile = strCommandLineArguments[3];

                    try { strSearch = strCommandLineArguments[4]; }
                    catch (Exception) { }
                };
            }
            else
            {
                // There is a user:
                boolRunAutomatically = false;
            };

            //------------------------------------------------------
            // Init all runtime criteria:

            // Data source:
            switch (strDataSource)
            {
                case "youtube":
                    radioButtonDataSrcYouTubeFavRef.Checked = true;
                    break;

                case "jimradio":
                    radioButtonDataSrcJimRadioRef.Checked = true;
                    break;

                case "import":
                    radioButtonDataSrcImportFileRef.Checked = true;
                    break;
            }

            // Channel:
            textBoxYoutubeFavoritesChannelRef.Text  = strChannel;

            // Search criteria (optional):
            textBoxSearchRef.Text = strSearch;

            // Sort:
            switch (strSort)
            {
                case "date":
                    radioButtonSortDateSavedRef.Checked = true;
                    break;

                case "title":
                    radioButtonSortTitleRef.Checked = true;
                    break;

                case "random":
                    radioButtonSortRandomRef.Checked = true;
                    break;
            }

            //------------------------------------------------------------------------
            if (boolRunAutomatically)
            {
                // Simulate the user's experience of clicking the buttons Retrieve,
                // Select All, and Play.  This fully automates the running of the
                // program.  NOTE: Since Play now launches an independent Windows/
                // Browser-based player, there is no longer any need for this form
                // and we can close it:
                Retrieve();

                if (intRowsDesired > 0)
                {
                    SelectAll();
                    Play();
                    formRef.Close();
                }
            };
            //------------------------------------------------------------------------

        } // core

        protected void Get_App_Config_Constants()
        {

            // Fill variables with values from App.Config making them, effectively, constants:
            YOUTUBE_API_BROWSER_KEY = Properties.Settings.Default.YOUTUBE_API_BROWSER_KEY;
            DATABASE_CONNECTION_STRING = Properties.Settings.Default.DATABASE_CONNECTION_STRING;

          //BROWSER = Properties.Settings.Default.BROWSER_CHROME;
          //BROWSER = Properties.Settings.Default.BROWSER_FIREFOX;
            BROWSER = Properties.Settings.Default.BROWSER_INTERNET_EXPLORER;
          //BROWSER = Properties.Settings.Default.BROWSER_OPERA;
          //BROWSER = Properties.Settings.Default.BROWSER_SAFARI;

            //----------------------------------------------------------------------
            // NOTE: I like to have the option to point to either the Production or
            // Development run environments:
            RUN_URL_PREFIX = Properties.Settings.Default.RUN_URL_PREFIX_PROD;
          //RUN_URL_PREFIX = Properties.Settings.Default.RUN_URL_PREFIX_DEV;
            //----------------------------------------------------------------------

        } // core

        // General purpose utility:
        public string get_delimited_substring(string strLookIn, string strDelimiter)
        {
            int intStartDelimiter = strLookIn.IndexOf(strDelimiter) + strDelimiter.Length;
            int intEndDelimiter   = strLookIn.IndexOf(strDelimiter, intStartDelimiter);

            return strLookIn.Substring(intStartDelimiter, intEndDelimiter - intStartDelimiter);
        }

        // NOTE: "async", "Task" and "await" are the hallmarks of asynchronous processing,
        // the only method by which these services are available.
        public async Task GetYoutubeServiceDeletionAuth()
        {
            UserCredential userCredential;

            // Get the "client_id" and "client_secret" codes, stored in a file:
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                userCredential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    // This OAuth 2.0 access scope allows an application to delete videos in the
                    // authorized user's YouTube channel.
                  //new[] { YouTubeService.Scope.YoutubeUpload },
                  //new[] { YouTubeService.Scope.Youtube },
                    new[] { YouTubeService.Scope.YoutubeForceSsl },
                    "user",
                    CancellationToken.None
                );
            }

            // This will be referenced later, in deletion processing:
            youTubeService = new YouTubeService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = userCredential,
                ApplicationName = Assembly.GetExecutingAssembly().GetName().Name
            });

        }

        //=========================================================================================      
        private int INIT_STEP_1_SignInAndGetYouTubeFavFeed()
        {
            //-------------------------------------------------------------------------------------------------
            // NOTE: This access, which is 1 of 2, provides a read-only view of any user's favorites playlist:
            // Supply registration and logon info to YouTube:

            // Fill variables with values from the GUI:
            YOUTUBE_ACCOUNT_CHANNEL   = textBoxYoutubeFavoritesChannelRef.Text;

            // Instantiate the service:
            youTubeService = new YouTubeService(new BaseClientService.Initializer()
            {
                ApiKey = YOUTUBE_API_BROWSER_KEY,
                ApplicationName = this.GetType().ToString()
            });

            // Access the channel's favorites playlist:
            var channelListRequest = youTubeService.Channels.List("contentDetails");
            channelListRequest.ForUsername = YOUTUBE_ACCOUNT_CHANNEL;

            // Retrieve the contentDetails part of the channel resource for the authorized user's channel.
            channelListResponse = channelListRequest.Execute();

            //-------------------------------------------------------------------------------------------------
            // NOTE: This access, which is 2 of 2, provides authorization necessary to delete videos from
            // a favorites playlist:

            int intReturnValue = 0; // init

            try
            {
                // Test the result by accessing it:
                strFavoritesListId = channelListResponse.Items[0].ContentDetails.RelatedPlaylists.Favorites;

                // Sign-in was successful:
                intReturnValue++;

                //- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                // Get the YouTube authorization necessary to delete videos from a favorites playlist:
                try
                {
                    // This processing fills variable "youTubeService":
                    GetYoutubeServiceDeletionAuth().Wait();

                    // Authorization was successful:
                    intReturnValue++;

                }
                catch (AggregateException ex)
                {
                    // Authorization was unsuccessful:
                    foreach (var e in ex.InnerExceptions)
                    {
                        MessageBox.Show(e.Message, "Authorization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                //- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

            }
            catch (Exception)
            {
                // Sign-in was unsuccessful:
            }

            // Possible return values:
            // 0 = sign-in was unsuccessful
            // 1 = sign-in was successful, authorization was unsuccessful
            // 2 = sign-in was successful, authorization was successful

            return intReturnValue;
            //-------------------------------------------------------------------------------------------------

        } // implementation: YouTube

        private void INIT_STEP_2_ShowInitSqlLoad()
        {
            // NOTE: This is the JimRadio hosted SQL Server database:

            // Init the SQL connection, statement and related objects:
            SqlConnection sqlConnection1 = new SqlConnection
            (
                Properties.Settings.Default.DATABASE_CONNECTION_STRING
            );

            SqlCommand sqlCommand1 = new SqlCommand("SELECT youtube_title, youtube_id, duration_seconds FROM video", sqlConnection1);
            SqlDataAdapter sqlDataAdapter1 = new SqlDataAdapter(sqlCommand1);
            DataSet dataSet1 = new DataSet();

            // Fill the DataSet with the results of the SQL query:
            sqlDataAdapter1.Fill(dataSet1);

            // Get the count of videos returned:
            intRowsRetrieved = dataSet1.Tables[0].Rows.Count;

            foreach (DataRow row in dataSet1.Tables[0].Rows)
            {
                // Init:
                bool boolShouldLoadEntry = true;

                // If the user specified search criteria:
                if (textBoxSearchRef.Text.Length > 0)
                {
                    // Determine if the title meets the search criteria:
                    if (TitleMeetsSearchCriteria(row["youtube_title"].ToString()) == false) { boolShouldLoadEntry = false; }
                };

                if (boolShouldLoadEntry)
                {
                    // Load the entry:
                    strInDateSavedOrigOrderTitle.Add(row["youtube_title"]);
                    strInDateSavedOrigOrderVideoId.Add(row["youtube_id"]);
                    strInDateSavedOrigOrderDuration.Add(row["duration_seconds"]);
                }

                // NOTE: For debugging:
                //MessageBox.Show(row["youtube_title"] + "\t" + row["youtube_id"] + "\t" + row["duration_seconds"]);
            }

        } // implementation: SQL

        private void INIT_STEP_2_ShowInitTextFileLoad()
        {

            // Init the get-filename dialog:
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog1.Title = "CSV text files";

            // Get/set the current runtime directory:
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();

            // If this value has not already been set by the command line:
            if (strImportFile == String.Empty)
            {
                // Get the file name for the import file:
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    strImportFile = openFileDialog1.FileName; 
                }
            }

            if (strImportFile != String.Empty)
            {
                if (File.Exists(strImportFile))
                {
                    intRowsRetrieved = 0; // init

                    // Open the file:
                    StreamReader streamReader1 = new StreamReader(strImportFile);

                    do
                    {
                        intRowsRetrieved++; // increment

                        // Read the contents of the current result as a comma-separated
                        // text file, using the StreamReader.

                        // Read a row from the file and break-out the columns:
                        string strFileRow = streamReader1.ReadLine();
                        string[] strColumn = strFileRow.Split(',');

                        // Init:
                        bool boolShouldLoadEntry = true;

                        // If the user specified search criteria:
                        if (textBoxSearchRef.Text.Length > 0)
                        {
                            // Determine if the title meets the search criteria:
                            if (TitleMeetsSearchCriteria(strColumn[0]) == false) boolShouldLoadEntry = false;
                        };

                        if (boolShouldLoadEntry)
                        {
                            // Load the entry:
                            strInDateSavedOrigOrderTitle.Add(strColumn[0]);    // title
                            strInDateSavedOrigOrderVideoId.Add(strColumn[1]);  // VideoId
                            strInDateSavedOrigOrderDuration.Add(strColumn[2]); // duration
                        }

                    } while (streamReader1.Peek() != -1);

                    // Close the file:
                    streamReader1.Close();
                }
            }

        } // implementation: CSV text file

        private void INIT_STEP_2_ShowInitYouTubeLoad()
        {
            //--------------------------------------------------------------------------------------------------------
            // SORT-ORDER 1 OF 3, Part A, by date saved (original):
            //--------------------------------------------------------------------------------------------------------

            intRowsRetrieved = 0; // init

            foreach (var channel in channelListResponse.Items)
            {
                // From the API response, extract the playlist ID that identifies the favorites list
                // of videos for the user's channel.
                var strFavoritesListId = channel.ContentDetails.RelatedPlaylists.Favorites;

                var nextPageToken = ""; // init
                while (nextPageToken != null)
                {
                    var playlistItemsListRequest = youTubeService.PlaylistItems.List("snippet");

                    playlistItemsListRequest.PlaylistId = strFavoritesListId;
                    playlistItemsListRequest.MaxResults = 50;
                    playlistItemsListRequest.PageToken = nextPageToken;

                    // Retrieve the list of videos saved to the favorites playlist:
                    var playlistItemsListResponse = playlistItemsListRequest.Execute();

                    // For each video:
                    foreach (var playlistItem in playlistItemsListResponse.Items)
                    {

                        intRowsRetrieved++; // increment

                        bool boolShouldLoadTitle = true; // init

                        // If the user specified search criteria:
                        if (textBoxSearchRef.Text.Length > 0)
                        {
                            if (TitleMeetsSearchCriteria(playlistItem.Snippet.Title) == false) { boolShouldLoadTitle = false; }
                        };

                        if (boolShouldLoadTitle)
                        {
                            // Load its info into an ArrayList, in original order.  But, before we
                            // do this, we will slightly edit the contents of the title:

                            // NOTE: A little cleaning-up goes a long way towards allowing the Title sort
                            // to collect like titles together:
                            strInDateSavedOrigOrderTitle.Add(AlterVideoTitle(playlistItem.Snippet.Title));
                            strInDateSavedOrigOrderVideoId.Add(playlistItem.Snippet.ResourceId.VideoId);
                            strInDateSavedOrigOrderItemId.Add(playlistItem.Id);

                            // NOTE: I would have to make a separate query just to get the video duration and I really
                            // don't need it so I won't bother:
                            strInDateSavedOrigOrderDuration.Add(Convert.ToInt32(0));

                        };

                    }

                    // Prep for the next time thru the loop:
                    nextPageToken = playlistItemsListResponse.NextPageToken;

                }
            }

        } // implementation: YouTube

        private void INIT_STEP_3_ShowSort()
        {
            //--------------------------------------------------------------------------------------------------------
            // SORT-ORDER 1 OF 3, Part B, by date saved (original):
            // Load the Favorites DataGridView directly (without a BindingSource):
            //--------------------------------------------------------------------------------------------------------
            if (radioButtonSortDateSavedRef.Checked == true)
            {
                // Set the sort mode:
                enumSortOrderSelected = EnumSortOrder.DateSaved; // init
                {
                    for (int i = 0; i < strInDateSavedOrigOrderTitle.Count; i++)
                    {
                        switch (enumDataSourceSelected)
                        {
                            case EnumDataSource.YouTubeFavorites:

                                dataGridViewRef.Rows.Add(strInDateSavedOrigOrderTitle[i],
                                                         strInDateSavedOrigOrderVideoId[i],
                                                         strInDateSavedOrigOrderDuration[i],
                                                         strInDateSavedOrigOrderItemId[i]
                                                        );
                                break;

                            case EnumDataSource.JimRadio:
                            case EnumDataSource.ImportFile:

                                dataGridViewRef.Rows.Add(strInDateSavedOrigOrderTitle[i],
                                                         strInDateSavedOrigOrderVideoId[i],
                                                         strInDateSavedOrigOrderDuration[i]
                                                        );
                                break;
                        }
                    }
                }
            }
            //--------------------------------------------------------------------------------------------------------
            // SORT-ORDER 2 OF 3, by Title:
            //-------------------------------------------------------------------------------------------------------
            else if (radioButtonSortTitleRef.Checked == true)
            {
                // Init:

                // Set the sort mode:
                enumSortOrderSelected = EnumSortOrder.Title;

                strInTitleOrderTitle = new ArrayList();
                strInTitleOrderVideoId = new ArrayList();
                strInTitleOrderDuration = new ArrayList();
                strInTitleOrderItemId = new ArrayList();

                // Init the Title Sort ArrayList by copying the contents of the Titles in their original order:
                strInTitleOrderTitle.AddRange(strInDateSavedOrigOrderTitle);

                // Now sort the Title ArrayList by title (using default Sort processing):
                strInTitleOrderTitle.Sort();

                // NOTE: We have sorted only the first, Title, of the three ArrayLists required.  Now
                // we must bring over the other two values, VideoId and Duration, so that all three
                // are a matching set:

                // Go thru each entry in the newly sorted ArrayList:
                for (int i = 0; i < strInTitleOrderTitle.Count; i++)
                {
                    // Init:

                    // Find what index position contained this Title in the original order:
                    int intOriginalPosition = strInDateSavedOrigOrderTitle.IndexOf(strInTitleOrderTitle[i]);

                    // Go to that position in the two other arrays of the set and bring their values over:
                    strInTitleOrderVideoId.Add(strInDateSavedOrigOrderVideoId[intOriginalPosition]);
                    strInTitleOrderDuration.Add(strInDateSavedOrigOrderDuration[intOriginalPosition]);

                    // Load the title to the GUI dataGridView:
                    switch (enumDataSourceSelected)
                    {
                        case EnumDataSource.YouTubeFavorites:

                            strInTitleOrderItemId.Add(strInDateSavedOrigOrderItemId[intOriginalPosition]);

                            dataGridViewRef.Rows.Add(strInTitleOrderTitle[i],
                                                     strInTitleOrderVideoId[i],
                                                     strInTitleOrderDuration[i],
                                                     strInTitleOrderItemId[i]
                                                    );
                            break;

                        case EnumDataSource.JimRadio:
                        case EnumDataSource.ImportFile:

                            dataGridViewRef.Rows.Add(strInTitleOrderTitle[i],
                                                     strInTitleOrderVideoId[i],
                                                     strInTitleOrderDuration[i]
                                                    );
                            break;

                    }
                }

                // NOTE: Row 1 is selected, so deselect it:
                dataGridViewRef.ClearSelection();
            }
            //--------------------------------------------------------------------------------------------------------
            // SORT-ORDER 3 OF 3, randomly:
            //--------------------------------------------------------------------------------------------------------
            else if (radioButtonSortRandomRef.Checked == true)
            {
                // Init:

                // Set the sort mode:
                enumSortOrderSelected = EnumSortOrder.Random;

                // Init the Title Sort ArrayList by copying the contents of the Titles in their original order.
                // I will work on a copy of the ArrayList because random sort processing empties out the (work)
                // ArrayList (ie, the copy):
                ArrayList strTitleInDateSavedOrigCopy = new ArrayList();
                strTitleInDateSavedOrigCopy.AddRange(strInDateSavedOrigOrderTitle);

                // Init:
                strInRandomOrderTitle = new ArrayList();
                strInRandomOrderVideoId = new ArrayList();
                strInRandomOrderDuration = new ArrayList();
                strInRandomOrderItemId = new ArrayList();

                //-----------------------------------------------------------------------------------------------
                // Begin sort process, driver array:

                // Use the random number generator:
                Random rnd = new Random(); // init

                // For all videos:

                // NOTE: this ArrayList's count will decrease as entries are taken from it:
                while (strTitleInDateSavedOrigCopy.Count > 0)
                {
                    // Pick a random item from the input list....
                    int intIndex = rnd.Next(0, strTitleInDateSavedOrigCopy.Count); // init

                    // .... and move it to the end of the output list:
                    strInRandomOrderTitle.Add(strTitleInDateSavedOrigCopy[intIndex]);
                    strTitleInDateSavedOrigCopy.RemoveAt(intIndex);
                }
                // ... at this point, the ArrayList is in a completely random order.

                // End sort process, driver array.
                //-----------------------------------------------------------------------------------------------
                // Begin sort process, other arrays:

                // NOTE: We have sorted only the first, Title, of the three ArrayLists required.  Now
                // we must bring over the other two values, VideoId and Duration, so that all three
                // are a matching set:

                // Go thru each entry in the newly sorted ArrayList:
                for (int i = 0; i < strInDateSavedOrigOrderTitle.Count; i++)
                {
                    // Find what index position contained this Title in the original order:
                    int intOriginalPosition = strInDateSavedOrigOrderTitle.IndexOf(strInRandomOrderTitle[i]); // init

                    // Go to that position in the two other arrays of the set and bring their values over:
                    strInRandomOrderVideoId.Add(strInDateSavedOrigOrderVideoId[intOriginalPosition]);
                    strInRandomOrderDuration.Add(strInDateSavedOrigOrderDuration[intOriginalPosition]);

                    // Load the title to the GUI dataGridView:
                    switch (enumDataSourceSelected)
                    {
                        case EnumDataSource.YouTubeFavorites:

                            strInRandomOrderItemId.Add(strInDateSavedOrigOrderItemId[intOriginalPosition]);

                            dataGridViewRef.Rows.Add(strInRandomOrderTitle[i],
                                                     strInRandomOrderVideoId[i],
                                                     strInRandomOrderDuration[i],
                                                     strInRandomOrderItemId[i]);
                            break;

                        case EnumDataSource.JimRadio:
                        case EnumDataSource.ImportFile:

                            dataGridViewRef.Rows.Add(strInRandomOrderTitle[i],
                                                     strInRandomOrderVideoId[i],
                                                     strInRandomOrderDuration[i]);
                            break;

                    }

                }
            // End sort process, other arrays.
            //---------------------------------------------------------------------------------------------------
        }
            //---------------------------------------------------------------------------------------------------

        } // core

        //=========================================================================================      

        public void Play() // PUBLIC  // core
        {
            PLAY_OR_DELETE_STEP_1_PlayOrDeleteUserSelections("play");
        }

        public void Delete() // PUBLIC  // core
        {
            PLAY_OR_DELETE_STEP_1_PlayOrDeleteUserSelections("delete");
        }

        //=========================================================================================      
        private void PLAY_OR_DELETE_STEP_1_PlayOrDeleteUserSelections(string strPlayOrDelete)
        {
            // Get the selected entries:
            int intPlayOrDeleteLoopSelectedRowCount =
                ((System.Windows.Forms.BaseCollection)(dataGridViewRef.SelectedRows)).Count;

            if (intPlayOrDeleteLoopSelectedRowCount > 0)
            {
                // Save the selected entries to their own array:
                int[] intPlayOrDeleteLoopSelectedRows = new int[intPlayOrDeleteLoopSelectedRowCount]; // init

                // Get a collection that contains the selected rows array, as rows:
                DataGridViewRow[] userSelectedRows = new DataGridViewRow[intPlayOrDeleteLoopSelectedRowCount]; // init
                dataGridViewRef.SelectedRows.CopyTo(userSelectedRows, 0);

                // NOTE: Array elements are in the reverse of the order selected.  Reverse them:
                Array.Reverse(userSelectedRows);

                // Loop thru the collection to determine, from the row structure definition, what was
                // the row number of the selected row:
                int intDestinationIndex = -1;
                foreach (DataGridViewRow dataGridViewRow in userSelectedRows)
                {
                    // Save these rows numbers in their own array:
                    intDestinationIndex++;
                    intPlayOrDeleteLoopSelectedRows[intDestinationIndex] = dataGridViewRow.Index;
                }

                switch (enumSortOrderSelected)
                {
                    //- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
                    case EnumSortOrder.DateSaved:

                        if (strPlayOrDelete == "play")
                        {
                            PLAY_OR_DELETE_STEP_2_PlayUserSelections(strInDateSavedOrigOrderVideoId, intPlayOrDeleteLoopSelectedRows);
                        }
                        else if (strPlayOrDelete == "delete")
                        {
                            PLAY_OR_DELETE_STEP_2_DeleteUserSelections(strInDateSavedOrigOrderItemId, intPlayOrDeleteLoopSelectedRows);
                        }

                        break;

                    //- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
                    case EnumSortOrder.Title:
                        if (strPlayOrDelete == "play")
                        {
                            PLAY_OR_DELETE_STEP_2_PlayUserSelections(strInTitleOrderVideoId, intPlayOrDeleteLoopSelectedRows);
                        }
                        else if (strPlayOrDelete == "delete")
                        {
                            PLAY_OR_DELETE_STEP_2_DeleteUserSelections(strInTitleOrderItemId, intPlayOrDeleteLoopSelectedRows);
                        }

                        break;

                    //- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
                    case EnumSortOrder.Random:
                        if (strPlayOrDelete == "play")
                        {
                            PLAY_OR_DELETE_STEP_2_PlayUserSelections(strInRandomOrderVideoId, intPlayOrDeleteLoopSelectedRows);
                        }
                        else if (strPlayOrDelete == "delete")
                        {
                            PLAY_OR_DELETE_STEP_2_DeleteUserSelections(strInRandomOrderItemId, intPlayOrDeleteLoopSelectedRows);
                        }

                        break;
                    //- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
                };
            }

        } // core

        private void PLAY_OR_DELETE_STEP_2_PlayUserSelections(
            ArrayList arrayListSortOrderVideoId, int[] intPlayOrDeleteLoopSelectedRows)
        {
            var URL_SIZE_BASED_MAXIMUM = 120; // NOTE: trial and error determined this number

            // Init:
            string strQuery = "";
            for (int i = 0; i < dataGridViewRef.SelectedRows.Count; i++)
            {
                // Since we are passing the selected videos by way of URL arguments,
                // we have to be aware of the maximum allowed size:
                if (i <= URL_SIZE_BASED_MAXIMUM)
                {
                    // Append:
                    if (i == 0) { strQuery += "?"; }
                    else { strQuery += "&"; };

                    // Append:
                    strQuery += "id=" + arrayListSortOrderVideoId[intPlayOrDeleteLoopSelectedRows[i]];
                }
            }

            // Play the selected videos:
            PLAY_STEP_3_PlayVideos(new Uri(RUN_URL_PREFIX + strQuery));

        } // core

        public void PLAY_OR_DELETE_STEP_2_DeleteUserSelections(
            ArrayList arrayListSortOrderItemId, int[] intPlayOrDeleteLoopSelectedRows)
        {
            string strPlaylistItemId;

            int intRowsDeleted = 0; // init

            // Confirm the deletion:
            if (MessageBox.Show("You have selected " + intPlayOrDeleteLoopSelectedRows.Length.ToString() + " videos for deletion.  Do you want to proceed?",
                                "Deletion Request", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                // Perform the deletion(s):
                for (int i = 0; i < dataGridViewRef.SelectedRows.Count; i++)
                {
                    strPlaylistItemId = arrayListSortOrderItemId[intPlayOrDeleteLoopSelectedRows[i]].ToString();

                    try
                    {
                        // Delete an individual video:
                        PlaylistItemsResource.DeleteRequest deleteRequest =
                            youTubeService.PlaylistItems.Delete(strPlaylistItemId);
                        deleteRequest.Execute();
                        intRowsDeleted++;
                    }
                    catch (Exception e)
                    {
                        // We will cancel the deletion process on a single error:
                        MessageBox.Show(get_delimited_substring(e.Message, "\r\n"),
                                        "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        break;
                    }

                }

                if (intRowsDeleted > 0)
                {
                    // Remove the out-of-date view:
                    dataGridViewRef.Rows.Clear();

                    // Inform the user about the results:
                    MessageBox.Show(intRowsDeleted.ToString() + " videos have been deleted.",
                                    "Deletion Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Given all the arrays that drive this process, it's far easier to just re-retrieve the results
                    // rather than deleting from the grid and synching up the arrays, etc:
                    Retrieve();
                }

            }

        } // core

        private void PLAY_STEP_3_PlayVideos(Uri urlToPlay) // core
        {
            // Instantiate a new browser Windows process...
            Process process = new Process();

            // ... that runs from this location...
            process.StartInfo.FileName = BROWSER;

            // ... and takes these arguments:
             process.StartInfo.Arguments = urlToPlay.ToString();           

            // Start it:
            process.Start();
        }

        //========================================================================================= 
        public void RegisterFormControls // PUBLIC  // external GUI interface
            (
              ref Form formRefArg, ref DataGridView dataGridViewArg, 

              ref GroupBox groupBoxDataSrcArg,
              ref RadioButton radioButtonDataSrcYouTubeFavArg, ref RadioButton radioButtonDataSrcJimRadioArg,
              ref RadioButton radioButtonDataSrcImportFileArg,            
            
              ref GroupBox groupBoxSortArg,
              ref RadioButton radioButtonSortDateSavedArg, ref RadioButton radioButtonSortTitleArg,
              ref RadioButton radioButtonSortRandomArg,

              ref TextBox textBoxYoutubeFavoritesChannelArg,
              ref TextBox textBoxSearchArg,
              ref Button buttonSignInAndRetrieveArg, ref Button buttonSelectAllArg,
              ref Button buttonPlayArg,
              ref Button buttonDeleteArg,
              ref Button buttonExportFileArg

            ) // core
        {
            // Save references:
            formRef = formRefArg;
            dataGridViewRef = dataGridViewArg;

            groupBoxDataSrcRef = groupBoxDataSrcArg;
            radioButtonDataSrcYouTubeFavRef = radioButtonDataSrcYouTubeFavArg;
            radioButtonDataSrcJimRadioRef = radioButtonDataSrcJimRadioArg;
            radioButtonDataSrcImportFileRef = radioButtonDataSrcImportFileArg;

            groupBoxSortRef = groupBoxSortArg;
            radioButtonSortDateSavedRef = radioButtonSortDateSavedArg;
            radioButtonSortTitleRef = radioButtonSortTitleArg;
            radioButtonSortRandomRef = radioButtonSortRandomArg;

            textBoxYoutubeFavoritesChannelRef  = textBoxYoutubeFavoritesChannelArg;
            textBoxSearchRef   = textBoxSearchArg;

            buttonSignInAndRetrieveRef = buttonSignInAndRetrieveArg;
            buttonSelectAllRef = buttonSelectAllArg;

            buttonPlayRef = buttonPlayArg;
            buttonDeleteRef = buttonDeleteArg;

            buttonExportFileRef = buttonExportFileArg;

        }

        public void Retrieve() // PUBLIC
        {
            // Get the Data Source:
            if (radioButtonDataSrcYouTubeFavRef.Checked == true)
            {
                enumDataSourceSelected = EnumDataSource.YouTubeFavorites;
            }
            else if (radioButtonDataSrcJimRadioRef.Checked == true)
            {
                enumDataSourceSelected = EnumDataSource.JimRadio;
            }
            else if (radioButtonDataSrcImportFileRef.Checked == true)
            {
                enumDataSourceSelected = EnumDataSource.ImportFile;
            }

            // Init:

            // NOTE: ArrayLists will be useful to us, later, as a
            // very handy way to do a random sort:
            strInDateSavedOrigOrderTitle = new ArrayList();
            strInDateSavedOrigOrderVideoId = new ArrayList();
            strInDateSavedOrigOrderDuration = new ArrayList();
            strInDateSavedOrigOrderItemId = new ArrayList();

            // Reset:
            dataGridViewRef.Rows.Clear();
            boolToggleSelectAll = false;
            strImportFile = "";

            string strMessage = "";

            // This column is  wider than the default:
            const int int_ITEM_ID_WIDTH = 410;

            // Reset if necessary:
            try
            {
                dataGridViewRef.Columns.Remove("Item Id");
            } catch (Exception) {}

            int intSignInAndauthorizationReturnValue = 0; // init

            // Based on the Data Source, run the appropriate retrieval and
            // begin to construct the window title:
            switch (enumDataSourceSelected)
            {
                case EnumDataSource.YouTubeFavorites:

                    // This data source requires an extra field:
                    int intColumn = dataGridViewRef.Columns.Add("Item Id", "");
                    dataGridViewRef.Columns[intColumn].Visible = false;

                    // Size the ItemId, which is very long, appropriately:
                    DataGridViewColumn DataGridViewColumn4 = dataGridViewRef.Columns[3];
                    DataGridViewColumn4.Width = int_ITEM_ID_WIDTH;

                    // Possible return values:
                    // 0 = sign-in was unsuccessful
                    // 1 = sign-in was successful, authorization was unsuccessful
                    // 2 = sign-in was successful, authorization was successful
                    intSignInAndauthorizationReturnValue = YouTubeSignInAndRetrieve();

                    strMessage = "YouTube Favorites - ";
                    break;

                case EnumDataSource.JimRadio:

                    SqlConnectAndRetrieve();
                    strMessage = "JimRadio - ";
                    break;

                case EnumDataSource.ImportFile:

                    TextFileImport();
                    strMessage = "Import File - ";
                    break;
            }

            // Get the count of videos desired (ie, that match the
            // search criteria):
            intRowsDesired = strInDateSavedOrigOrderTitle.Count;

            // Construct the message...
            strMessage += Convert.ToString(intRowsDesired) + " of " +
                          Convert.ToString(intRowsRetrieved);             

            // ... and put it in the window title:
            SetFormTitle(strMessage);

            // NOTE: With a zero row count, the Select, Play, etc. buttons should
            // not be enabled:
            if (intRowsDesired > 0)
            {
                // We are now on this step and in this mode:
                SetFormControlsModally(EnumFormMode.step3SignedInWithList);

                // By default, the first row is selected:
                dataGridViewRef.ClearSelection();

                // Set delete access:
                switch (enumDataSourceSelected)
                {
                    case EnumDataSource.YouTubeFavorites:

                        // 2 = sign-in was successful, authorization was successful
                        if (intSignInAndauthorizationReturnValue == 2)
                        {
                            // We are now on this step and in this mode:
                            SetFormControlsModally(EnumFormMode.step4AllowDelete);
                        }

                        break;

                    case EnumDataSource.JimRadio:
                    case EnumDataSource.ImportFile:

                        // We are now on this step and in this mode:
                        SetFormControlsModally(EnumFormMode.step4PreventDelete);
                        break;

                }

            }

        } // implementation: driver

        public void SelectAll() // PUBLIC
        {
            // Toggle the boolean by NOT-ing it:
            boolToggleSelectAll = !boolToggleSelectAll;

            if (boolToggleSelectAll)
            {
                dataGridViewRef.SelectAll();
            }
            else
            {
                dataGridViewRef.ClearSelection();
            }

        } // core

        private static void SetFormControlsModally(EnumFormMode enumFormModeCurrent)
        {
            switch (enumFormModeCurrent)
            {
                // NOTE: Step 1 is conceptual but is no different than Step 2:
                case EnumFormMode.step1UninitializedForm:
                case EnumFormMode.step2InitializedForm:

                    groupBoxSortRef.Enabled = true;
                    groupBoxDataSrcRef.Enabled = true;

                    buttonSignInAndRetrieveRef.Enabled = true;
                    textBoxYoutubeFavoritesChannelRef.Enabled = true;
                    textBoxSearchRef.Enabled = true;
                    dataGridViewRef.Enabled = true;

                    buttonPlayRef.Enabled = false;
                    buttonDeleteRef.Enabled = false;
                    buttonSelectAllRef.Enabled = false;

                    buttonExportFileRef.Enabled = false;

                    break;                   

                case EnumFormMode.step3SignedInWithList:

                    groupBoxSortRef.Enabled = true;
                    groupBoxDataSrcRef.Enabled = true;

                    buttonSignInAndRetrieveRef.Enabled = true;
                    textBoxYoutubeFavoritesChannelRef.Enabled = true;
                    textBoxSearchRef.Enabled = true;
                    dataGridViewRef.Enabled = true;

                    buttonPlayRef.Enabled = true;
                    buttonSelectAllRef.Enabled = true;

                    buttonExportFileRef.Enabled = true;

                    break;

                case EnumFormMode.step4AllowDelete:

                    buttonDeleteRef.Enabled = true;

                    break;

                case EnumFormMode.step4PreventDelete:

                    buttonDeleteRef.Enabled = false;

                    break;
            }

        } // core

        private void SetFormTitle(string strMessage)
        {
            formRef.Text = strMessage;

        } // core

        public void SqlConnectAndRetrieve()
        {
            Cursor.Current = Cursors.WaitCursor;

            // Load from SQL:
            INIT_STEP_2_ShowInitSqlLoad();
            INIT_STEP_3_ShowSort();

            Cursor.Current = Cursors.Default;

        } // implementation: SQL    

        public void TextFileImport()
        {
            Cursor.Current = Cursors.WaitCursor;

            // Load from CSV text file:
            INIT_STEP_2_ShowInitTextFileLoad();
            INIT_STEP_3_ShowSort();

            Cursor.Current = Cursors.Default;

        } // implementation: CSV text file

        private bool ThisCriterionMatchesThisTitle(string strCriterion, string strTitle)
        {
            if (strTitle.ToLower().IndexOf(strCriterion.ToLower()) > -1) { return true; }
            else { return false; };

        } // core

        private bool TitleMeetsSearchCriteria(string strTitle)
        {
            // Get the comma-separated list of search criteria from the GUI:
            string strSearchCriteria = textBoxSearchRef.Text;

            // Break it out into an array:
            string[] strSearchCriteriaElements = strSearchCriteria.Split(new Char[] { ',' });

            // Init to "no match":
            Boolean boolCriterionMatchesTitle = false;

            foreach (string strElement in strSearchCriteriaElements)
            {
                // Compare each of the search criteria elements to the current Title:
                boolCriterionMatchesTitle = ThisCriterionMatchesThisTitle(strElement.Trim(), strTitle);

                // If there was a match:
                if (boolCriterionMatchesTitle)
                {
                    // We can stop searching:
                    break;
                }
            }

            return boolCriterionMatchesTitle;

        } // core

        private int YouTubeSignInAndRetrieve() // PUBLIC
        {
            // Possible return values:
            // 0 = sign-in was unsuccessful
            // 1 = sign-in was successful, authorization was unsuccessful
            // 2 = sign-in was successful, authorization was successful

            int intReturnValue = INIT_STEP_1_SignInAndGetYouTubeFavFeed();

            if (intReturnValue > 0)
            {
                Cursor.Current = Cursors.WaitCursor;

                // Load from YouTube:
                INIT_STEP_2_ShowInitYouTubeLoad();
                INIT_STEP_3_ShowSort();

                Cursor.Current = Cursors.Default;
            }
            else
            {
                MessageBox.Show("Sign-in failed", "YouTube Sign-In",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            return intReturnValue;

        } // implementation: YouTube

        //-----------------------------------------------------------------------------
        // END METHODS

    }

 }
