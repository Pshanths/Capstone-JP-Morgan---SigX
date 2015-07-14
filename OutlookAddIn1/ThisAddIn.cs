//Invoke all required libraries 
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Data.SqlClient;
using OnenoteOCR;
using System.Drawing;
using System.Diagnostics;
using System.Text.RegularExpressions;

//Project Namespace
namespace OutlookAddIn1
{
    //Main Class 
    public partial class ThisAddIn
    {
        //Global Class Variables
        //Outlook.NameSpace outlookNameSpace;
        public static string[] data = new string[10]; // String Array to store data(fields) before inserting into database
        public static bool skipTextReader; // Boolean value to determine if OCR should be run, or not
        public static Stopwatch timer = new Stopwatch(); // Stop watch object to measure total time for parsing
        public static double time_elapsed; // Time taken to parse emails 
        public static bool dont_add; 
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
            

        //Method that runs when the plugin loads 
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Prompt to run application, and start and end timer
            System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("Do you want to run SigX?", "Hit yes, I'll be quick", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                timer = Stopwatch.StartNew(); ReadMail();  //Start stopwatch and run batch parsing function
            }
            //Configure plugin to run when email is added to Inbox 
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);
            items = inbox.Items;

            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(ReadSingleMail); // Modified method to run for single email


        }


        //Method that runs on a single email
        static void ReadSingleMail(dynamic item)
        {
            
            timer = Stopwatch.StartNew(); 
            //Reinitialize data array back to empty
            for (int index = 0; index <= 8; index++)
            {
                data[index] = "";
            }
            skipTextReader = false; // Reinitialize variable back to false

            //Set up OCR engine to prepare for OCR
            OnenoteOcrEngine ocr = new OnenoteOcrEngine();

            
            string bodyText; // Email body
            string mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //Path to My Documents
            StringBuilder sb = new StringBuilder();

            if (ReadImageFromEmail(item) == 1) // If email has an attachment, read the attachment- Assuming this is E-business card
            {
                Image imageToParse = Image.FromFile(@"C:\TestFileSave\" + item.Attachments[1].FileName);

                //Get text from OCR, and normalize 
                bodyText = ocr.Recognize(imageToParse);
                bodyText = bodyText.Replace(" ", ";"); bodyText = bodyText.Replace(";;", ";");
                bodyText = bodyText.Replace(",", ""); bodyText = bodyText.Replace("|", "");

                skipTextReader = true; //Skip parsing email body if email has image signature


                data = parseOCRSignature(bodyText); // Parse body into fields 
            }

            else //If no attachment, extract email body
            {
                if (item != null)
                {
                    bodyText = item.Body;
                }
                else
                {
                    return; // If no email body, exit function.
                }
                sb.Append(bodyText);
            }


            if (skipTextReader == false) //If no image signature, parse email body
            {
                File.AppendAllText(mydocpath + @"\MailFile.txt", sb.ToString());
                sb.Remove(0, bodyText.Length);
                string signature = extractSignature().Replace(";;", ";"); //Extract signature block
                //IF EMBEDDED IMAGE SIGNATURE FOUND
                if (((signature.ToLower()).Contains(".png") || (signature.ToLower()).Contains(".gif") || (signature.ToLower()).Contains(".jpg")) && signature.Length < 20)
                {
                    //MessageBox.Show(signature);
                    Image imageToParse = Image.FromFile(@"C:\TestFileSave\" + signature);
                    bodyText = ocr.Recognize(imageToParse);
                    //MessageBox.Show(bodyText);
                    if (bodyText != null)
                    {
                        bodyText = bodyText.Replace(" ", ";");
                        bodyText = bodyText.Replace(";;", ";");
                        bodyText = bodyText.Replace(",", "");
                        bodyText = bodyText.Replace("|", ";");
                        skipTextReader = true; //Bool variable if email has image signature
                        data = parseOCRSignature(bodyText);
                    }
                }
                else
                {
                    data = parseImageSignature(signature); //Parse signature block     
                }
            }

            //Send data to database
            string strEmailSenderEmailId = Convert.ToString(item.SenderEmailAddress);
            SqlConnection connection = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=C:\Users\Prashanth\Desktop\CAPSTONE JPMORGAN\OutlookAddIn1\OutlookContacts.mdf;Integrated Security=True;User Instance=True");
            SqlCommand command = new SqlCommand();
            SqlDataReader dataReader;
            command.Connection = connection;
            connection.Open();
            command.CommandText = "select * from New_Contacts where Email_ID = '" + strEmailSenderEmailId + "'";
            dataReader = command.ExecuteReader();

            //Check if contact exists already
            if (dataReader.HasRows == false)
            {
                command.CommandText = "insert into New_Contacts VALUES('" + data[0] + "' , '" + data[1] + "' , '" + data[2] + "' , '" + data[3] + "' , '" + data[4] + "' , '" + data[5] + "' , '" + data[6] + "' , '" + data[7] + "' , '" + data[8] + "' , '" + strEmailSenderEmailId + "' , '" + data[9] + "')";
                dataReader.Close();
                command.ExecuteNonQuery();
            }
            //If it does, update and overwrite selected data. This method prevents overwriting past data with empty fields
            else
            {
                UpdateSelectFields(data[2], data[4], data[7], data[5], data[6], data[8], data[9], strEmailSenderEmailId);

            }

            // Delete the file, stop the timer
            System.IO.File.Delete(mydocpath + @"\MailFile.txt");
            timer.Stop();
            time_elapsed = time_elapsed + timer.Elapsed.TotalSeconds;
            ocr.Dispose();

            MessageBox.Show("This email took " + (time_elapsed).ToString() + " seconds!");

        }

        // This method prevents overwriting past data with empty fields. Assumes Name, and Company doesn't change for a given emailID
        static void UpdateSelectFields(string title, string address, string cell, string home, string fax, string Website, string sig_EmailID, string emailID)
        {
            //Set up DB connections
            SqlConnection connection = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=C:\Users\Prashanth\Desktop\CAPSTONE JPMORGAN\OutlookAddIn1\OutlookContacts.mdf;Integrated Security=True;User Instance=True");
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            connection.Open();

            //For each field, check if the new values are empty or NULL, if not update existing DB values
            if (!(string.IsNullOrWhiteSpace(title)))
            {
                command.CommandText = "update New_Contacts SET title = '" + title + "' WHERE Email_ID = '" + emailID + "'";
                command.ExecuteNonQuery();
            }

            if (!(string.IsNullOrWhiteSpace(address)))
            {
                command.CommandText = "update New_Contacts SET address = '" + address + "' WHERE Email_ID = '" + emailID + "'";
                command.ExecuteNonQuery();
            }

            if (!(string.IsNullOrWhiteSpace(cell)))
            {
                command.CommandText = "update New_Contacts SET cell_phone = '" + cell + "' WHERE Email_ID = '" + emailID + "'";
                command.ExecuteNonQuery();
            }

            if (!(string.IsNullOrWhiteSpace(fax)))
            {
                command.CommandText = "update New_Contacts SET Fax  = '" + fax + "' WHERE Email_ID = '" + emailID + "'";
                command.ExecuteNonQuery();
            }

            if (!(string.IsNullOrWhiteSpace(home)))
            {
                command.CommandText = "update New_Contacts SET Tel_phone = '" + home + "' WHERE Email_ID = '" + emailID + "'";
                command.ExecuteNonQuery();
            }

            if (!(string.IsNullOrWhiteSpace(Website)))
            {
                command.CommandText = "update New_Contacts SET Website = '" + Website + "' WHERE Email_ID = '" + emailID + "'";
                command.ExecuteNonQuery();
            }

            if (!(string.IsNullOrWhiteSpace(sig_EmailID)))
            {
                command.CommandText = "update New_Contacts SET sig_Email_id = '" + sig_EmailID + "' WHERE Email_ID = '" + emailID + "'";
                command.ExecuteNonQuery();
            }
           

        }
        

        //Method that returns 1 if there are attachments in the email
        static int ReadImageFromEmail(Outlook.MailItem item)
        {
            if (item != null)
            {
                if (item.Attachments.Count > 0)
                {
                    for (int i = 1; i <= item
                       .Attachments.Count; i++)
                       {
                            // Save only image attachments 
                           if ((item.Attachments[i].FileName.ToLower().EndsWith("png")) || item.Attachments[i].FileName.ToLower().EndsWith("jpg") || item.Attachments[i].FileName.ToLower().EndsWith("jpeg") || item.Attachments[i].FileName.ToLower().EndsWith("gif"))
                           {
                               item.Attachments[i].SaveAsFile
                                   (@"C:\TestFileSave\" +
                                   item.Attachments[i].FileName);
                           }
                    }
                    return 1;
                }
            }
            return 0;
        }


        //Big Brother of ReadSingleMail() - Deals with email batches - Runs only once when Outlook is first launched
        static void ReadMail() 
        {
            

            //Set up OCR
            OnenoteOcrEngine ocr = new OnenoteOcrEngine();

            string bodyText;
            string mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            //Get unread emails from Inbox
            Microsoft.Office.Interop.Outlook.Application app = null;
            Microsoft.Office.Interop.Outlook._NameSpace ns = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder inboxFolder = null;
            app = new Microsoft.Office.Interop.Outlook.Application();
            ns = app.GetNamespace("MAPI");
            inboxFolder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items unreadItems = inboxFolder.Items.Restrict("[Unread]=true");
            int max_runs;
            //Go through each Unread email
            if (unreadItems.Count > 10) { max_runs = 10; }
            else max_runs = unreadItems.Count;
            
            for (int counter = 1; counter <= max_runs; counter++)
            {
                //Reinitialize Data array
                for (int index = 0; index <= 8; index++)
                {
                    data[index] = "";
                }
                skipTextReader = false;

                dynamic item = unreadItems[counter];
                StringBuilder sb = new StringBuilder();

                if (ReadImageFromEmail(item) == 1) // If email has an attachment, read the attachment
                {
                    Image imageToParse = Image.FromFile(@"C:\TestFileSave\" + item.Attachments[1].FileName);
                    bodyText = ocr.Recognize(imageToParse);
                    //MessageBox.Show(bodyText);
                    if (bodyText != null)
                    {
                        bodyText = bodyText.Replace(" ", ";");
                        bodyText = bodyText.Replace(";;", ";");
                        bodyText = bodyText.Replace(",", ""); 
                        bodyText = bodyText.Replace("|", ";");
                        skipTextReader = true; //Bool variable if email has image signature
                        
                        data = parseOCRSignature(bodyText);
                        
                    }
                    else
                    {
                        break;
                    }
                }
                
                else //If no attachment, extract email body
                {
                    if (item != null)
                    {
                        bodyText = item.Body;
                    }
                    else
                    {
                        continue;
                    }
                    sb.Append(bodyText);
                }
                

                if (skipTextReader == false) //If no image signature, parse email body
                {
                    File.AppendAllText(mydocpath + @"\MailFile.txt", sb.ToString());
                    sb.Remove(0, bodyText.Length);
                    string signature = extractSignature().Replace(";;", ";"); //Extract signature block
                    //IF EMBEDDED IMAGE SIGNATURE FOUND
                    if (((signature.ToLower()).Contains(".png") || (signature.ToLower()).Contains(".gif") || (signature.ToLower()).Contains(".jpg")) && signature.Length<20)
                    {
                        //MessageBox.Show(signature);
                        Image imageToParse = Image.FromFile(@"C:\TestFileSave\" + signature);
                        bodyText = ocr.Recognize(imageToParse);
                        //MessageBox.Show(bodyText);
                        if (bodyText != null)
                        {
                            bodyText = bodyText.Replace(" ", ";");
                            bodyText = bodyText.Replace(";;", ";");
                            bodyText = bodyText.Replace(",", "");
                            bodyText = bodyText.Replace("|", ";");
                            skipTextReader = true; //Bool variable if email has image signature
                            data = parseOCRSignature(bodyText);
                        }
                    }
                    else
                    {
                        data = parseImageSignature(signature); //Parse signature block     
                    }     
                }

                //Send data to database
                //Get Sender Email Address
                string strEmailSenderEmailId = Convert.ToString(item.SenderEmailAddress);

                //Set up DB connections
                SqlConnection connection = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=C:\Users\Prashanth\Desktop\CAPSTONE JPMORGAN\OutlookAddIn1\OutlookContacts.mdf;Integrated Security=True;User Instance=True");
                SqlCommand command = new SqlCommand();
                SqlDataReader dataReader;
                command.Connection = connection;
                connection.Open();


                command.CommandText = "select * from New_Contacts where Email_ID = '" + strEmailSenderEmailId + "'";
                dataReader = command.ExecuteReader();

                //Check if contact exists already
                if (dataReader.HasRows == false)
                {
                    command.CommandText = "insert into New_Contacts VALUES('" + data[0] + "' , '" + data[1] + "' , '" + data[2] + "' , '" + data[3] + "' , '" + data[4] + "' , '" + data[5] + "' , '" + data[6] + "' , '" + data[7] + "' , '" + data[8] + "' , '" + strEmailSenderEmailId + "' , '" + data[9] + "')";
                    dataReader.Close();
                    command.ExecuteNonQuery();
                }
                //If it does, update and overwrite past data
                else
                {
                    UpdateSelectFields(data[2], data[4], data[7], data[5], data[6], data[8] , data[9], strEmailSenderEmailId);

                }
               
            }

            System.IO.File.Delete(mydocpath + @"\MailFile.txt");
            timer.Stop();
            time_elapsed = time_elapsed + timer.Elapsed.TotalSeconds;
            ocr.Dispose();
            
            MessageBox.Show("SigX took " + (time_elapsed).ToString() + " seconds for " + unreadItems.Count +
                "emails!");
            //connection.Close();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            MessageBox.Show("Finished");
        }

        //Method to eliminate Hyperlinks from numbers
        public static String cleanString(String s)
        {
            s.Trim();
            s = s.Replace("hyperlink", "");
            s = s.Replace(" ", "");
            //finds the last index where ' occurs
            int end = s.LastIndexOf("\"");
            int num = s.Length;

            int n = s.Length - end;

            char[] a = s.ToCharArray();
            char[] b = new char[256];

            int j = 0;
            for (int i = end + 1; i < s.Length; i++)
            {

                b[j] = a[i];
                j++;
            }

            string str = new string(b);
            return str;
        }


        //Method that seperates last name and company seperated by a semicolon
        static string extractLastName(string lastnameAndCompany)
        {
            //MessageBox.Show(lastnameAndCompany);
            string lastName = "";
            int index = 0;
            while (lastnameAndCompany[index] != ';')
            {
                lastName = lastName + lastnameAndCompany[index];
                index++;
            }

            return lastName;
        }

        //Method that locates the signature in the Email Body
        static string extractSignature() //FUNCTION TO EXTRACT SIGNATURE BLOCK FROM EMAIL
        {
            string mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            using (StreamReader reader = new StreamReader(mydocpath + @"\MailFile.txt"))
            {
                string line;
                int i = -1;
                int n = 0;
                int c = 0;
                string[] siglines = new string[500];

                //READ EMAIL AND STOP AT END OF SIGNATURE BLOCK WHILE STORING BODY INTO ARRAY OF STRINGS BY LINE
                while ((line = reader.ReadLine()) != null)
                {
                    i++;
                    siglines[i] = line;
                    if ((line.ToLower()).StartsWith("this email is confidential") || (line.ToLower()).StartsWith("this email and any files") || (line.ToLower().Contains("original message")))
                    {
                        i--;
                        break;
                    }
                    else if((line.ToLower()).StartsWith(">"))
                    {
                        if (c==1)
                        {
                            if ((line.ToLower()).StartsWith(">"))
                            {
                                i=i-2;
                                break;
                            }
                        }
                        c++;
                    }
                }

                // NOW READ THROUGH ARRAY OF STIRNGS BACKWARDS TO DETERMINE WHEN SIGNATURE BLOCK STARTS		
                for (n = i; n > 0; n--)
                {
                    
                    //Check for ascii lines, excluding the very bottom of the email
                    if (n < i && (siglines[n].Contains("--") || siglines[n].Contains("___") || siglines[n].Contains("~~~") || siglines[n].Contains("***") || siglines[n].Contains("^^^") || siglines[n].Contains("---")))
                    {
                        break; //(break out of for loop)
                    }
                    //Check for key words that dictate end of main body of email
                    if ((siglines[n].ToLower()).Contains("thank") || (siglines[n].ToLower()).Contains("best") || (siglines[n].ToLower()).Contains("regards") || (siglines[n].ToLower()).Contains("sincere")
                      || (siglines[n].ToLower()).Contains("rgds") || (siglines[n].ToLower()).Contains("hope") || (siglines[n].ToLower()).Contains("take care") || (siglines[n].ToLower()).Contains("thx")
                      || (siglines[n].ToLower()).Contains("peace") || (siglines[n].ToLower()).Contains("cheers") || (siglines[n].ToLower()).Contains("yours") || (siglines[n].ToLower()).Contains("love")
                      || (siglines[n].ToLower()).Contains("see you") || (siglines[n].ToLower()).Contains("bye") || (siglines[n].ToLower()).Contains("bless") || (siglines[n].ToLower()).Contains("happy")
                      || (siglines[n].ToLower()).Contains("sent from my") || (siglines[n].ToLower()).Contains("warm") || (siglines[n].ToLower()).Contains("look forward"))
                    {
                        //do not count this line and everything below is signature
                        n = n + 1;
                        break; //(break out of for loop)
                    }
                }
                
               
                //Verify length of extracted signature
                if (n == 0) //Extracted the whole email
                {

                    dont_add = true;
                    //There is no signature
                    //MessageBox.Show("NO SIGNATURE DETECTED");
                }

                else if (i - n > 20) //Extracted block greater than 10 lines (extracted body is double spaced hence 20)
                {
                    
                    //MessageBox.Show("Signature over 10 lines");
                }
               
                else if (n == i-1) //Extracted signature is only one line
                {
                    //Check if more than 3 characters to exclude informal or initialized signatures
                    if (siglines[n].Length <= 3)
                    {
                        //YOU PROBABLY ONLY GOT INITIALS, DISREGARD SIGNATURE AND MOVE ON
                        //MessageBox.Show("Signature under 4 characters");
                    }
                   
                }

                //Convert siglines array into a single signature string and return
                string signature = string.Join(";", siglines, n, (i - n + 1));
                //Check if embedded image signature
                if (!((signature.ToLower()).Contains("\r\n") || ((signature.ToLower()).Contains("\n"))))
                {
                  if (((signature.ToLower()).Contains("http")) && ((signature.ToLower()).Contains(".jpg") || (signature.ToLower()).Contains(".png") || (signature.ToLower()).Contains(".gif")))
                      {
                        WebClient webClient = new WebClient();
                        int end = signature.LastIndexOf("/");
                        string imgname = signature.Remove(0, end + 1);
                        signature = signature.Replace(";", "");
                        signature = signature.Trim();
                        //MessageBox.Show(signature);
                        //MessageBox.Show(imgname);
                        webClient.DownloadFile(signature, @"C:\TestFileSave\" + imgname);
                        webClient.Dispose();
                        return imgname;
                      }
                }
                reader.Close();
                return signature;
            }
        }


        public static String NumberscleanString(String s)
        {

            //finds the last index where ' occurs
            int end = s.LastIndexOf("\"");


            string str = s.Remove(0, end + 1);
            //MessageBox.Show(str);
            return str;
        }


        static string[] parseImageSignature(string signaturetext) //FUNCTION TO PARSE EXTRACTED SIGNATURE BLOCK
        {

            
            //Array to store fields
            string[] fields = new string[10];
            string AddressString = "";

            //Checker Variable for firstname and lastname
            int checker = 0;

            //Normalize Signature Text
            signaturetext = signaturetext.ToLower();
            signaturetext = signaturetext.Replace("|", ";");
            //MessageBox.Show(signaturetext);
            //MessageBox.Show(signaturetext);
            string[] words = signaturetext.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);// Splits the signature and saves each word as an element

            //Go through each word in the signature 
            foreach (string s in words)
            {
                //MessageBox.Show(s);
                //Check to see if there any empty entries 
                string test = s.Trim();
                if (string.IsNullOrEmpty(test) || (string.IsNullOrWhiteSpace(test))) continue;

                // FIRST NAME & LAST NAME

                if (checker == 0) //We are looking for names
                {
                  
                    string[] names;
                    string temp;

                    if (skipTextReader == false) //Split First Line of Signature 
                    {
                        //Remove any starting semicolons
                        if (s.StartsWith(";"))
                            temp = s.TrimStart(';');

                        else
                            temp = s;
                        //MessageBox.Show("temp is " + temp);
                        
                        //Names array is split into as many words as there are names
                        names = temp.Split(new char[] {' '} , StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < names.Length; i++)
                        {
                            names[i].Trim();

                        }
                    }


                    else
                    {
                        names = signaturetext.Split(';');
                    }


                    if (names.Length == 1)//Only one name. Put it in first Name
                    {
                        fields[0] = names[0];
                         //Assume there was no last name
                    }
                   
                    else if (names.Length == 0) // No Name??
                    {
                        checker++; // Dont look at names again after this
                        continue;
                    
                    }

                    else //First name, Middle name, Second Middle name, Third Middle name, ....... Last name
                    {
                        fields[0] = names[0]; //First word is assumed first name
                        fields[1] = names[names.Length - 1]; // Last word is assumed last name
                    }
                    checker++; //Never look for names after this
                    continue;
                }



                //POSITION
                //Check if word matches certain keywords
                else if (test.ToLower().Contains("vp") || test.ToLower().Contains("banker") || test.ToLower().Contains("vice") || test.ToLower().Contains("president") || test.ToLower().Contains("chief") || test.ToLower().Contains("manager") || test.ToLower().Contains("developer") || test.ToLower().Contains("lead") || test.ToLower().Contains("advisor") || test.ToLower().Contains("partner")
                || test.ToLower().Contains("designer") || (test.ToLower().Contains("senior")) || test.ToLower().Contains("junior") || test.ToLower().Contains("analyst") || test.ToLower().Contains("consultant") || test.ToLower().Contains("architect") || test.ToLower().Contains("supervis") || test.ToLower().Contains("secretary") || test.ToLower().Contains("associate") || test.ToLower().Contains("operator")
                || test.ToLower().Contains("official") || test.ToLower().Contains("assistant") || test.ToLower().Contains("director") || test.ToLower().Contains("engineer") || test.ToLower().Contains("quality") || test.ToLower().Contains("CEO") || test.ToLower().Contains("CFO") || test.ToLower().Contains("rep") || test.ToLower().Contains("specialist") || test.ToLower().Contains("captain") || test.ToLower().Contains("student"))
                {
                    fields[2] = s; continue;
                }



                //ADDRESS
                //Address modified
                else if ((test.Any(char.IsDigit)) && ((test.Contains(",")) || (test.Contains("st") || (test.Contains("dr")) || (test.Contains("ave")) || (test.Contains("blvd")) || (test.Contains("boulevard")) || (test.Contains("rd")) || (test.Contains("rt")) || (test.Contains("route")) || (test.Contains("road")) || (test.Contains("hwy")) || (test.Contains("highway")))))
                {
                    AddressString  += test;

                    fields[4] = AddressString;

                }

                //NUMBERS
                else if (test.Any(char.IsDigit))
                {

                    string test2 = test.Replace("-", string.Empty);
                    test2 = test2.Replace("(", string.Empty);
                    test2 = test2.Replace(")", string.Empty);
                    test2 = test2.Replace(" ", string.Empty);
                    test2 = test2.Replace(".", string.Empty);
                    test2 = test2.Replace(":", string.Empty);
                    //MessageBox.Show(test2);

                    if (test2.All(char.IsDigit) && (test2.Length == 10))
                    {
                        test2 = test2.Insert(0, "(");
                        test2 = test2.Insert(4, ") ");
                        test2 = test2.Insert(9, "-");

                        fields[7] = test2;
                    }

                    else if (test2.ToLower().Contains("tel"))
                    {


                        test2 = test2.Replace("tel", string.Empty);

                        if (test2.ToLower().Contains("t") || test2.ToLower().Contains("home") || test2.ToLower().Contains("work") || test2.ToLower().Contains("w"))
                        {
                            test = NumberscleanString(test);
                            test2 = test.Replace("-", string.Empty);
                            test2 = test2.Replace("(", string.Empty);
                            test2 = test2.Replace(")", string.Empty);
                            test2 = test2.Replace(" ", string.Empty);
                            test2 = test2.Replace(".", string.Empty);
                            test2 = test2.Replace(":", string.Empty);
                            if (test2.Length == 12)
                            {
                                test2 = test2.Insert(2, "(");
                                test2 = test2.Insert(6, ") ");
                                test2 = test2.Insert(11, "-");
                            }
                            else if (test2.Length == 11)
                            {
                                test2 = test2.Insert(1, "(");
                                test2 = test2.Insert(5, ") ");
                                test2 = test2.Insert(10, "-");
                            }
                            else
                            {
                                test2 = test2.Insert(0, "(");
                                test2 = test2.Insert(4, ") ");
                                test2 = test2.Insert(9, "-");
                            }
                            //MessageBox.Show(test2);
                            fields[5] = test2;
                        }

                        else if (test.ToLower().Contains("f") || test.ToLower().Contains("fax"))
                        {
                            test = NumberscleanString(test);
                            test2 = test.Replace("-", string.Empty);
                            test2 = test2.Replace("(", string.Empty);
                            test2 = test2.Replace(")", string.Empty);
                            test2 = test2.Replace(" ", string.Empty);
                            test2 = test2.Replace(".", string.Empty);
                            test2 = test2.Replace(":", string.Empty);
                            if (test2.Length == 12)
                            {
                                test2 = test2.Insert(2, "(");
                                test2 = test2.Insert(6, ") ");
                                test2 = test2.Insert(11, "-");
                            }
                            else if (test2.Length == 11)
                            {
                                test2 = test2.Insert(1, "(");
                                test2 = test2.Insert(5, ") ");
                                test2 = test2.Insert(10, "-");
                            }
                            else
                            {
                                test2 = test2.Insert(0, "(");
                                test2 = test2.Insert(4, ") ");
                                test2 = test2.Insert(9, "-");
                            }
                            //MessageBox.Show(test2);
                            fields[6] = test2;
                        }

                        else if (test2.ToLower().Contains("c") || test2.ToLower().Contains("cell") || test2.ToLower().Contains("mobile") || test2.ToLower().Contains("m"))
                        {
                            test = NumberscleanString(test);
                            test2 = test.Replace("-", string.Empty);
                            test2 = test2.Replace("(", string.Empty);
                            test2 = test2.Replace(")", string.Empty);
                            test2 = test2.Replace(" ", string.Empty);
                            test2 = test2.Replace(".", string.Empty);
                            test2 = test2.Replace(":", string.Empty);
                            if (test2.Length == 12)
                            {
                                test2 = test2.Insert(2, "(");
                                test2 = test2.Insert(6, ") ");
                                test2 = test2.Insert(11, "-");
                            }
                            else if (test2.Length == 11)
                            {
                                test2 = test2.Insert(1, "(");
                                test2 = test2.Insert(5, ") ");
                                test2 = test2.Insert(10, "-");
                            }
                            else
                            {
                                test2 = test2.Insert(0, "(");
                                test2 = test2.Insert(4, ") ");
                                test2 = test2.Insert(9, "-");
                            }
                            //MessageBox.Show(test2);
                            fields[7] = test2;
                        }
                    }
                    else if (test2.ToLower().Contains("t") || test2.ToLower().Contains("home") || test2.ToLower().Contains("work") || test2.ToLower().Contains("w"))
                    {

                        test2 = Regex.Replace(test2, "[^0-9]", "");
                        if (test2.Length != 10)
                        {
                            break;
                        }
                        test2 = test2.Insert(0, "(");
                        test2 = test2.Insert(4, ") ");
                        test2 = test2.Insert(9, "-");

                        fields[5] = test2;
                    }

                    else if (test.ToLower().Contains("f") || test.ToLower().Contains("fax"))
                    {
                        test2 = Regex.Replace(test2, "[^0-9]", "");
                        if (test2.Length != 10)
                        {
                            break;
                        }

                        test2 = test2.Insert(0, "(");
                        test2 = test2.Insert(4, ") ");
                        test2 = test2.Insert(9, "-");

                        fields[6] = test2;
                    }

                    else if (test2.ToLower().Contains("c") || test2.ToLower().Contains("cell") || test2.ToLower().Contains("mobile") || test2.ToLower().Contains("m"))
                    {

                        test2 = Regex.Replace(test2, "[^0-9]", "");
                        if (test2.Length != 10)
                        {
                            break;
                        }
                        test2 = test2.Insert(0, "(");
                        test2 = test2.Insert(4, ") ");
                        test2 = test2.Insert(9, "-");

                        fields[7] = test2;
                    }

                }

                // WEBSITE AND SIGNATURE EMAILID              
                else if (((test.EndsWith("com")) || (test.EndsWith("edu"))))
                {
                    if (test.Contains("@"))
                    {

                        fields[9] = cleanString(test);
                        continue;
                    }

                    else
                    {

                        fields[8] = cleanString(test);
                        continue;

                    }
                }

                //COMPANY
                // Queries 'Comapnies' table to see if word is a company. If not look for inc, LLC
                SqlConnection connection = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=C:\Users\Prashanth\Desktop\CAPSTONE JPMORGAN\OutlookAddIn1\OutlookContacts.mdf;Integrated Security=True;User Instance=True");
                SqlCommand command = new SqlCommand();
                SqlDataReader dataReader;
                command.Connection = connection;
                connection.Open();

                test = test.Replace(".", string.Empty); test = test.Trim(); test = test.Replace(" ", string.Empty);
                command.CommandText = "select company_name from Companies where company_name LIKE '%" + test + "%'";
                //command.CommandText = "select company_name from Companies where company_name = '" + test + "'";
                dataReader = command.ExecuteReader();

                if ((dataReader.HasRows) && !(string.IsNullOrWhiteSpace(s)))
                {
                    fields[3] = s;
                    continue;
                }

                else if (test.Contains("llc") || test.Contains("inc"))
                {
                    command.CommandText = "INSERT INTO COMPANIES values('" + test + "')";
                    dataReader.Close();
                    command.ExecuteNonQuery();
                    fields[3] = s;
                    continue;
                }
                //connection.Close();
            }


            //MessageBox.Show("First Name: " + fields[0]);
            // MessageBox.Show("Last Name: " + fields[1]);
             //MessageBox.Show("Position: " + fields[2]);
             //MessageBox.Show("Company: " + fields[3]);
             //MessageBox.Show("Address: " + fields[4]);
             //MessageBox.Show("Home Number: " + fields[5]);
             //MessageBox.Show("Fax Number: " + fields[6]);
             //MessageBox.Show("Cell Number: " + fields[7]);
             //MessageBox.Show("Website: " + fields[8]);
            return fields;
        }
        static string[] parseOCRSignature(string signaturetext) //FUNCTION TO PARSE OCR IMAGE
        {
           

            //Array to store fields
            string[] fields = new string[10];
            string AddressString = "";

            //Checker Variable for firstname and lastname
            int checker = 0;
            string[] names;
             string temp;
            //Normalize Signature Text
            signaturetext = signaturetext.ToLower();
            signaturetext = signaturetext.Replace("|", ";");
            signaturetext = signaturetext.Replace(";", " ");
           // MessageBox.Show(signaturetext);
            string[] words = signaturetext.Split(new string[] { Environment.NewLine, "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);// Splits the signature and saves each word as an element
            
            
            //Go through each word in the signature 
            foreach (string s in words)
            {
                
                //Check to see if there any empty or irrelevant entries 
                string test = s.Trim();
                
                if (string.IsNullOrEmpty(test) || (string.IsNullOrWhiteSpace(test))) continue;
                if ((s.ToLower()).Contains("thank") || (s.ToLower()).Contains("best") || (s.ToLower()).Contains("regards") || (s.ToLower()).Contains("sincere")
                      || (s.ToLower()).Contains("rgds") || (s.ToLower()).Contains("hope") || (s.ToLower()).Contains("take care") || (s.ToLower()).Contains("thx")
                      || (s.ToLower()).Contains("peace") || (s.ToLower()).Contains("cheers") || (s.ToLower()).Contains("yours") || (s.ToLower()).Contains("love")
                      || (s.ToLower()).Contains("see you") || (s.ToLower()).Contains("bye") || (s.ToLower()).Contains("bless") || (s.ToLower()).Contains("happy")
                      || (s.ToLower()).Contains("sent from my") || (s.ToLower()).Contains("warm") || (s.ToLower()).Contains("look forward"))
                    {
                        continue;
                    }
                if ((s.ToLower()).Contains("footer") || (s.ToLower()).Contains("authorize") || (s.ToLower()).Contains("register") || (s.ToLower()).Contains("authorise"))
                {
                    continue;
                }

                // FIRST NAME & LAST NAME
                if (checker == 0) //We are looking for names
                {
                //Remove any starting semicolons
                        if (s.StartsWith(";"))
                            temp = s.TrimStart(';');

                        else
                            temp = s;
                        //MessageBox.Show("names are: " + temp);
                        
                        //Names array is split into as many words as there are names
                        names = temp.Split(new char[] {' '} , StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < names.Length; i++)
                        {
                            names[i].Trim();

                        }
        


                    if (names.Length == 1)//Only one name. Put it in first Name
                    {
                        fields[0] = names[0];
                         //Assume there was no last name
                    }
                   
                    else if (names.Length == 0) // No Name??
                    {
                        checker++; // Dont look at names again after this
                        continue;
                    
                    }

                    else //First name, Middle name, Second Middle name, Third Middle name, ....... Last name
                    {
                        fields[0] = names[0]; //First word is assumed first name
                        fields[1] = names[names.Length - 1]; // Last word is assumed last name
                    }
                    checker++; //Never look for names after this
                    continue;
                }



                //POSITION
                //Check if word matches certain keywords
                else if (test.ToLower().Contains("vp") || test.ToLower().Contains("banker") || test.ToLower().Contains("vice") || test.ToLower().Contains("president") || test.ToLower().Contains("chief") || test.ToLower().Contains("manager") || test.ToLower().Contains("developer") || test.ToLower().Contains("lead") || test.ToLower().Contains("advisor") || test.ToLower().Contains("partner")
                || test.ToLower().Contains("designer") || (test.ToLower().Contains("senior")) || test.ToLower().Contains("junior") || test.ToLower().Contains("analyst") || test.ToLower().Contains("consultant") || test.ToLower().Contains("architect") || test.ToLower().Contains("supervis") || test.ToLower().Contains("secretary") || test.ToLower().Contains("associate") || test.ToLower().Contains("operator")
                || test.ToLower().Contains("official") || test.ToLower().Contains("assistant") || test.ToLower().Contains("director") || test.ToLower().Contains("engineer") || test.ToLower().Contains("quality") || test.ToLower().Contains("CEO") || test.ToLower().Contains("CFO") || test.ToLower().Contains("rep") || test.ToLower().Contains("specialist") || test.ToLower().Contains("captain") || test.ToLower().Contains("student"))
                {
                    fields[2] = s; continue;
                }



                //ADDRESS
                //Address modified
                else if ((test.Any(char.IsDigit)) && ((test.Contains(",")) || (test.Contains("st") || (test.Contains("dr")) || (test.Contains("ave")) || (test.Contains("blvd")) || (test.Contains("boulevard")) || (test.Contains("rd")) || (test.Contains("rt")) || (test.Contains("route")) || (test.Contains("road")) || (test.Contains("hwy")) || (test.Contains("highway")))))
                {
                    AddressString += test;

                    fields[4] = AddressString;

                }

                //NUMBERS
                else if (test.Any(char.IsDigit))
                {

                    string test2 = test.Replace("-", string.Empty);
                    test2 = test2.Replace("(", string.Empty);
                    test2 = test2.Replace(")", string.Empty);
                    test2 = test2.Replace(" ", string.Empty);
                    test2 = test2.Replace(".", string.Empty);
                    test2 = test2.Replace(":", string.Empty);
                    //MessageBox.Show("unformatted num: " + test2);

                    if (test2.All(char.IsDigit) && (test2.Length == 10))
                    {
                        test2 = test2.Insert(0, "(");
                        test2 = test2.Insert(4, ") ");
                        test2 = test2.Insert(9, "-");
                        fields[7] = test2;
                    }

                   if (test2.ToLower().Contains("t") || test2.ToLower().Contains("home") || test2.ToLower().Contains("work") || test2.ToLower().Contains("w"))
                    {

                        test2 = Regex.Replace(test2, "[^0-9]", "");
                        if (test2.Length == 10)
                        {
                            test2 = test2.Insert(0, "(");
                            test2 = test2.Insert(4, ") ");
                            test2 = test2.Insert(9, "-");

                            fields[5] = test2;
                        }
                    }

                    else if (test.ToLower().Contains("f") || test.ToLower().Contains("fax"))
                    {
                        test2 = Regex.Replace(test2, "[^0-9]", "");
                        if (test2.Length == 10)
                        {
                            test2 = test2.Insert(0, "(");
                            test2 = test2.Insert(4, ") ");
                            test2 = test2.Insert(9, "-");

                            fields[6] = test2;
                        }
                    }

                    else if (test2.ToLower().Contains("c") || test2.ToLower().Contains("cell") || test2.ToLower().Contains("mobile") || test2.ToLower().Contains("m"))
                    {
                        test2 = Regex.Replace(test2, "[^0-9]", "");
                        if (test2.Length == 10)
                        {
                            test2 = test2.Insert(0, "(");
                            test2 = test2.Insert(4, ") ");
                            test2 = test2.Insert(9, "-");

                            fields[7] = test2;
                        }
                    }

                }

                // WEBSITE AND SIGNATURE EMAILID              
                else if ((test.EndsWith(".com")) || (test.EndsWith(".edu")) || (test.EndsWith(".gov")) || (test.EndsWith(".org")) || (test.Contains(".co")))
                {
                    if (test.Contains("@"))
                    {

                        fields[9] = cleanString(test);
                        continue;
                    }

                    else
                    {

                        fields[8] = cleanString(test);
                        continue;

                    }
                }

                //COMPANY
                // Queries 'Comapnies' table to see if word is a company. If not look for inc, LLC
                SqlConnection connection = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=C:\Users\Prashanth\Desktop\CAPSTONE JPMORGAN\OutlookAddIn1\OutlookContacts.mdf;Integrated Security=True;User Instance=True");
                SqlCommand command = new SqlCommand();
                SqlDataReader dataReader;
                command.Connection = connection;
                connection.Open();

                test = test.Replace(".", string.Empty); test = test.Trim(); test = test.Replace(" ", string.Empty);
                if(test.StartsWith("@"))
                {
                    test.Remove(0, 1);
                }
                command.CommandText = "select company_name from Companies where company_name LIKE '%" + test + "%'";
                //command.CommandText = "select company_name from Companies where company_name = '" + test + "'";
                dataReader = command.ExecuteReader();

                if ((dataReader.HasRows) && !(string.IsNullOrWhiteSpace(s)))
                {
                    fields[3] = s;
                    continue;
                }

                else if (test.Contains("llc") || test.Contains("inc"))
                {
                    command.CommandText = "INSERT INTO COMPANIES values('" + test + "')";
                    dataReader.Close();
                    command.ExecuteNonQuery();
                    fields[3] = s;
                    continue;
                }
                //connection.Close();
            }


            //MessageBox.Show("First Name: " + fields[0]);
            // MessageBox.Show("Last Name: " + fields[1]);
            //MessageBox.Show("Position: " + fields[2]);
            //MessageBox.Show("Company: " + fields[3]);
            //MessageBox.Show("Address: " + fields[4]);
            //MessageBox.Show("Home Number: " + fields[5]);
            //MessageBox.Show("Fax Number: " + fields[6]);
            //MessageBox.Show("Cell Number: " + fields[7]);
            //MessageBox.Show("Website: " + fields[8]);
            return fields;
        }


        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
    }
}


