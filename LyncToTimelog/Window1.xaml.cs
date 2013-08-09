using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Windows;
using System.Xml;

namespace LyncToTimelog
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        private const string ApplicationGuid = "{FA44026B-CC48-42DA-AAA8-B849BCB12345}";

        public XmlDocument contactsRaw { get; set; }
        public XmlDocument customersRaw { get; set; }
        private static readonly DirectoryInfo AppPath = new FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).Directory;

        public Window1()
        {
            try
            {
                string url = ConfigurationManager.AppSettings["timelogPath"];
                string siteCode = ConfigurationManager.AppSettings["siteCode"];
                string apiId = ConfigurationManager.AppSettings["apiId"];
                string apiPassword = ConfigurationManager.AppSettings["apiPassword"];

                InitializeComponent();
                this.Hide();

                LyncClient.GetClient().ConversationManager.ConversationAdded += ConversationManager_ConversationAdded;
                LyncClient.GetClient().ConversationManager.ConversationRemoved += ConversationManager_ConversationRemoved;

                var _endpoint = new EndpointAddress(url + "/service.asmx");
                var _binding = new BasicHttpBinding();
                _binding.MaxReceivedMessageSize = 100000000;

                if (url.Contains("https"))
                {
                    _binding.Security.Mode = BasicHttpSecurityMode.Transport;
                }

                TimelogApi.ServiceSoapClient timelogClient = new TimelogApi.ServiceSoapClient(_binding, _endpoint);

                if (!File.Exists(AppPath + "\\contactsRaw.xml"))
                {
                    File.AppendAllText(AppPath + "\\contactsRaw.xml", new XmlDocument().CreateXmlDeclaration("1.0", "UTF-8", null).OuterXml);
                    File.AppendAllText(AppPath + "\\contactsRaw.xml", timelogClient.GetContactsRaw(siteCode, apiId, apiPassword, 0, 0).OuterXml);
                }

                if (!File.Exists(AppPath + "\\customersRaw.xml"))
                {
                    File.AppendAllText(AppPath + "\\customersRaw.xml", new XmlDocument().CreateXmlDeclaration("1.0", "UTF-8", null).OuterXml);
                    File.AppendAllText(AppPath + "\\customersRaw.xml", timelogClient.GetCustomersRaw(siteCode, apiId, apiPassword, 0, -1, 0, string.Empty).OuterXml);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured while finding the Lync Client and fetching data from TimeLog\r\n\r\n" + ex.Message, "LyncToTimelog - Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void ConversationManager_ConversationRemoved(object sender, ConversationManagerEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Action)(() =>
            {
                this.Topmost = false;
                this.Hide();
                this.WindowState = System.Windows.WindowState.Minimized;
            }));
        }

        void ConversationManager_ConversationAdded(object sender, Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs e)
        {
            var window = LyncClient.GetAutomation().GetConversationWindow(e.Conversation);

            //e.Conversation.Modalities[ModalityTypes.AudioVideo].ModalityStateChanged

            var top = window.Top;
            var left = window.Left;
            var width = window.Width;
            var height = window.Height;

            var selfuri = e.Conversation.SelfParticipant.Contact.Uri;
            var uri = e.Conversation.Participants.Where(c => c.Contact.Uri != selfuri).Select(c => c.Contact.Uri).FirstOrDefault();

            this.Dispatcher.BeginInvoke((Action)(() =>
                {
                    this.LabelParticipantName.Content = "Loading...";
                    this.LabelNumbersList.Content = "Loading...";
                    this.LabelAddressName.Content = "Loading...";
                    this.LabelCompanyName.Content = "Loading...";
                    this.LabelEmailName.Content = "Loading...";

                    this.WindowState = System.Windows.WindowState.Normal;
                    this.Topmost = true;
                    this.Top = System.Windows.SystemParameters.PrimaryScreenHeight - this.Height - 50;
                    this.Left = System.Windows.SystemParameters.PrimaryScreenWidth - this.Width - 10;

                    this.Show();

                    var contact = LyncClient.GetClient().ContactManager.GetContactByUri(uri);
                    var endpoints = contact.GetContactInformation(ContactInformationType.ContactEndpoints);
                    var displayName = contact.GetContactInformation(ContactInformationType.DisplayName);

                    this.LabelParticipantName.Content = displayName;

                    var customersRaw = new XmlDocument();
                    customersRaw.Load(AppPath + "\\customersRaw.xml");
                    var customerManager = new XmlNamespaceManager(customersRaw.NameTable);
                    customerManager.AddNamespace("tlp", "http://www.timelog.com/XML/Schema/tlp/v4_4");

                    var contactsRaw = new XmlDocument();
                    contactsRaw.Load(AppPath + "\\contactsRaw.xml");
                    var contactManager = new XmlNamespaceManager(contactsRaw.NameTable);
                    contactManager.AddNamespace("tlp", "http://www.timelog.com/XML/Schema/tlp/v4_4");

                    // List phone numbers
                    this.LabelNumbersList.Content = string.Empty;
                    foreach (var item in (endpoints as IList<Object>))
                    {
                        var endpoint = (item as ContactEndpoint);
                        string endpointuri = endpoint.Uri.Replace("tel:", string.Empty).Replace("sip:", string.Empty);

                        if (endpoint.Type == ContactEndpointType.WorkPhone || endpoint.Type == ContactEndpointType.MobilePhone)
                        {
                            this.LabelNumbersList.Content += endpointuri + "\r\n";
                        }
                    }

                    // Loop phone numbers to match with TimeLog Project
                    bool isFound = false;
                    foreach (var item in (endpoints as IList<Object>))
                    {
                        if (isFound)
                        {
                            break;
                        }

                        string endpointuri = (item as ContactEndpoint).Uri.Replace("tel:", string.Empty).Replace("sip:", string.Empty);
                        foreach (XmlNode contactNode in contactsRaw.SelectSingleNode("tlp:Contacts", contactManager).ChildNodes)
                        {
                            if ((contactNode.SelectSingleNode("tlp:Phone", contactManager) != null && contactNode.SelectSingleNode("tlp:Phone", contactManager).InnerText == endpointuri) ||
                                (contactNode.SelectSingleNode("tlp:Mobile", contactManager) != null && contactNode.SelectSingleNode("tlp:Mobile", contactManager).InnerText == endpointuri))
                            {
                                this.LabelAddressName.Content = contactNode.SelectSingleNode("tlp:Address1", contactManager).InnerText + "\r\n";
                                this.LabelAddressName.Content += contactNode.SelectSingleNode("tlp:ZipCode", contactManager).InnerText;
                                this.LabelAddressName.Content += contactNode.SelectSingleNode("tlp:City", contactManager).InnerText;
                                this.LabelCompanyName.Content = contactNode.SelectSingleNode("tlp:CustomerName", contactManager).InnerText;
                                this.LabelEmailName.Content = contactNode.SelectSingleNode("tlp:Email", contactManager).InnerText;

                                isFound = true;
                                break;
                            }
                        }
                    }

                    if (!isFound)
                    {
                        this.LabelAddressName.Content = "Not found";
                        this.LabelCompanyName.Content = "Not found";
                        this.LabelEmailName.Content = "Not found";
                    }
                }));
        }
    }
}