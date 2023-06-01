using System.Collections;

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using System.Xml.Linq;
using System.Collections.Specialized;


namespace EmailSpammer3._0
{
    public partial class Form1 : Form
    {
        public ArrayList PasswordsList = new ArrayList();
        private int emailIndex;
        private string Attachment1;
        private string Attachment2;
        private string Attachment3;
        private string Attachment4;
        private string Attachment5;

        public Form1()
        {
            InitializeComponent();
            

        }

        private void emailadd_utton_Click(object sender, EventArgs e)
        {
            if(textBox1.Text.Contains("@"))
            SpamTargets.Items.Add(textBox1.Text);
            textBox1.Text= "";
        }

        private void SpamTargets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(SpamTargets.SelectedIndex != -1)
            {
                string TitleItem = SpamTargets.SelectedItem.ToString();
                DialogResult dialogResult = MessageBox.Show("Do you want to remove : " + TitleItem, "Remove Item", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    SpamTargets.Items.Remove(SpamTargets.SelectedItem);
                }

            }

        }

        private void SpamSenders_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SpamSenders.SelectedIndex != -1)
            {
                string TitleItem = SpamSenders.SelectedItem.ToString();
                DialogResult dialogResult = MessageBox.Show("Do you want to remove : " + TitleItem, "Remove Item", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    SpamSenders.Items.Remove(SpamSenders.SelectedItem);
                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Contains("@"))
            {
                SpamSenders.Items.Add(textBox2.Text);
                PasswordsList.Add(textBox3.Text);
            }
                
            textBox2.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Create an instance of OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the filter to only allow image files
            openFileDialog.Filter = "Image Files (*.webp;*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp";

            // Show the dialog and check if the user selected a file
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                // Find the first unused Attachment variable and assign the attachment
                if (Attachment1 == null)
                {
                    Attachment1 = filePath;
                    pictureBox1.Image = Image.FromFile(filePath);
                }
                else if (Attachment2 == null)
                {
                    Attachment2 = filePath;
                    pictureBox2.Image = Image.FromFile(filePath);
                }
                else if (Attachment3 == null)
                {
                    Attachment3 = filePath;
                    pictureBox3.Image = Image.FromFile(filePath);
                }
                else if (Attachment4 == null)
                {
                    Attachment4 = filePath;
                    pictureBox4.Image = Image.FromFile(filePath);
                }
                else if (Attachment5 == null)
                {
                    Attachment5 = filePath;
                    pictureBox5.Image = Image.FromFile(filePath);
                }
                else
                {
                    MessageBox.Show("All attachment slots are filled.");
                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Attachment1 = null;
            pictureBox1.Image = null;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Attachment2 = null;
            pictureBox2.Image = null;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Attachment4 = null;
            pictureBox4.Image = null;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            Attachment3 = null;
            pictureBox3.Image = null;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            Attachment5 = null;
            pictureBox5.Image = null;
        }

        private void Emails_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            button3.Text = "Spamming";
            emailIndex = 0;
            int amount;
            string[] spamTargetsArray = SpamTargets.Items.Cast<string>().ToArray();

            var tasks = new List<Task>();
            if (int.TryParse(amountTextBox.Text, out amount))
            {
                for (int i = 1; i <= amount; i++)
                {
                    for (int j = 0; j < SpamSenders.Items.Count; j++)
                    {
                        var fromEmail = SpamSenders.Items[j].ToString();
                        var fromPassword = PasswordsList[j].ToString();

                        
                        tasks.Add(Task.Run(() => SendEmailOutlook(fromEmail, fromPassword, spamTargetsArray, subjectTextbox.Text, descTextBox.Text)));
                        
                        EmailSent();
                    }
                }
            }

            Task.WhenAll(tasks).Wait();
            button3.Text = "Start";
            button3.Enabled = true;
        }

        private void SendEmailOutlook(string fromEmail, string fromPassword, string[] toEmailArray, string subject, string body)
        {
            

            try
            {
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("", fromEmail));
                foreach (string toEmail in toEmailArray)
                {
                    message.To.Add(new MailboxAddress(toEmail.Trim(), toEmail.Trim()));
                }
                emailIndex++;
                message.Subject = subject + "(" + emailIndex + ")";
                var builder = new BodyBuilder();
                builder.HtmlBody = body;
                if (Attachment1 != null)
                {
                    builder.Attachments.Add(Attachment1);
                }
                if (Attachment2 != null)
                {
                    builder.Attachments.Add(Attachment2);
                }
                if (Attachment3 != null)
                {
                    builder.Attachments.Add(Attachment3);
                }
                if (Attachment4 != null)
                {
                    builder.Attachments.Add(Attachment4);
                }
                if (Attachment5 != null)
                {
                    builder.Attachments.Add(Attachment5);
                }
                message.Body = builder.ToMessageBody();

               

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.office365.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(fromEmail, fromPassword);
                    client.Send(message);
                    client.Disconnect(true);

                    

                }

            }
            catch (Exception ex)
            {
                
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error sending email retrying...");
                    Console.ForegroundColor = ConsoleColor.Green;
                emailIndex--;
                    SendEmailOutlook(fromEmail, fromPassword, toEmailArray, subject, body);
                
               
            }
        }

        private void EmailSent()
        {
            Emails.Items.Add("Email sending..");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ;
            // Clear existing items in comboBox1 (optional)
            comboBox1.Items.Clear();

            // Retrieve the savenames from the SavedData property in Properties.Settings
            if (Properties.Settings.Default.SavedData != null)
            {
                var savedData = Properties.Settings.Default.SavedData;
                comboBox1.Items.AddRange(savedData.Cast<string>().ToArray());
            }

            


        }

        private void button4_Click(object sender, EventArgs e)
        {
            string saveName = comboBox1.Text;

            // Create a new dictionary to store the saved data
            Dictionary<string, List<string>> savedData = new Dictionary<string, List<string>>();

            // Retrieve the existing saved data from Properties.Settings.Default.SavedData
            if (Properties.Settings.Default.SavedData != null)
            {
                // Convert the existing StringCollection to Dictionary<string, List<string>>
                savedData = ConvertToDictionary(Properties.Settings.Default.SavedData);
            }

            // Add or update the saved data for the specified saveName
            List<string> saveDataList = new List<string>();

            // Retrieve all items from the senderemailsListBox and add them to saveDataList
            foreach (var item in SpamSenders.Items)
            {
                saveDataList.Add(item.ToString());
            }

            // Retrieve all items from the targetemailsListBox and add them to saveDataList
            foreach (var item in SpamTargets.Items)
            {
                saveDataList.Add(item.ToString());
            }

            // Retrieve all items from the passwordsListBox and add them to saveDataList
            foreach (var item in PasswordsList)
            {
                saveDataList.Add(item.ToString());
            }

            savedData[saveName] = saveDataList;

            // Convert the updated saved data back to StringCollection
            Properties.Settings.Default.SavedData = ConvertToStringCollection(savedData);

            // Save the changes
            Properties.Settings.Default.Save();
        }

        // Helper method to convert Dictionary<string, List<string>> to StringCollection
        private StringCollection ConvertToStringCollection(Dictionary<string, List<string>> dictionary)
        {
            StringCollection stringCollection = new StringCollection();

            foreach (var pair in dictionary)
            {
                stringCollection.Add(pair.Key); // Add the saveName
                foreach (var item in pair.Value)
                {
                    stringCollection.Add(item); // Add each item in the list
                }
            }

            return stringCollection;
        }

        // Helper method to convert StringCollection to Dictionary<string, List<string>>
        private Dictionary<string, List<string>> ConvertToDictionary(StringCollection stringCollection)
        {
            Dictionary<string, List<string>> dictionary = new Dictionary<string, List<string>>();
            List<string> currentList = null;

            foreach (var item in stringCollection)
            {
                string str = item.ToString();

                if (!string.IsNullOrEmpty(str))
                {
                    if (!dictionary.ContainsKey(str))
                    {
                        dictionary[str] = new List<string>();
                        currentList = dictionary[str];
                    }
                }
                else
                {
                    if (currentList != null)
                    {
                        currentList.Add(str);
                    }
                }
            }

            return dictionary;
        }



        private void button5_Click(object sender, EventArgs e)
        {
            string selectedSavename = comboBox1.Text;

            // Retrieve the savedData from the SavedData property in Properties.Settings
            if (Properties.Settings.Default.SavedData != null)
            {
                var savedData = Properties.Settings.Default.SavedData;

                // Check if the selected savename exists in the savedData
                if (savedData.Contains(selectedSavename))
                {
                    // Retrieve the index of the selected savename
                    int selectedIndex = savedData.IndexOf(selectedSavename);

                    // Calculate the start index of the associated passwords, senders, and targets
                    int passwordIndex = selectedIndex;
                    int senderIndex = selectedIndex + 1;
                    int targetIndex = selectedIndex + 2;

                    // Retrieve the passwords, senders, and targets associated with the selected savename
                    string[] passwords = savedData[passwordIndex + 1].Split(';');
                    string[] senders = savedData[senderIndex + 1].Split(';');
                    string[] targets = savedData[targetIndex + 1].Split(';');

                    // Clear the existing items in the list boxes
                    SpamSenders.Items.Clear();
                    SpamTargets.Items.Clear();
                    PasswordsList.Clear();

                    // Add items from passwords array to PasswordsList
                    foreach (var password in passwords)
                    {
                        PasswordsList.Add(password);
                    }

                    // Add items from senders array to SpamSenders list box
                    foreach (var spamer in senders)
                    {
                        SpamSenders.Items.Add(spamer);
                    }

                    // Add items from targets array to SpamTargets list box
                    foreach (var target in targets)
                    {
                        SpamTargets.Items.Add(target);
                    }


                    // Use the passwords, senders, and targets as needed
                    // ...
                }
                else
                {
                    // Handle the case when the selected savename doesn't exist
                    // ...
                }
            }
        }


    }
}