using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Autodiscover;
using NDesk.Options;

namespace gal_dump
{
    class Program
    {
        private static readonly string[] StartArray = new string[]
        {
            "a", "b", "c", "d", "e", "f", "g", "h", "i",
            "j", "k", "l", "m", "n", "o", "p", "q", "r",
            "s", "t", "u", "v", "w", "x", "y", "z", "_",
            "@", ".", "-"
        };

        public static List<AddressClass> AllContacts = new List<AddressClass>();
        public static List<String> AllFoundGroupNames = new List<String>();

        private static Boolean ContactExists(String emailAddress)
        {
            try
            {
                if (AllContacts.Count.Equals(0))
                {
                    return false;
                }

                foreach (AddressClass thisAddress in AllContacts)
                {
                    if (thisAddress.Address.Equals(emailAddress))
                    {
                        return true;
                    }
                }
            }
            catch (NullReferenceException ex)
            {
                return false;
            }

            return false;
        }

        private static Object GetValue(Contact thisContact, PropertyDefinitionBase schemaObject)
        {
            Object returnValue = null;

            try
            {
                thisContact.TryGetProperty(schemaObject, out returnValue);
            }
            catch (Exception ex)
            {
            }

            return returnValue;
        }

        private static String ObjectToString(Object inValue)
        {
            String returnValue;
            try
            {
                returnValue = inValue.ToString();
            }
            catch (NullReferenceException ex)
            {
                return String.Empty;
            }

            return returnValue;
        }

        private static String GetPhone(Contact thisContact, PhoneNumberKey inKey)
        {
            String phone;
            try
            {
                phone = thisContact.PhoneNumbers[inKey];
            }
            catch (KeyNotFoundException ex)
            {
                phone = "";
            }

            return phone;
        }

        private static String GetEmail(Contact thisContact, EmailAddressKey inKey)
        {
            String email;

            try
            {
                email = thisContact.EmailAddresses[inKey].Address;
                email = email.Replace("SMTP:", "");
            }
            catch (KeyNotFoundException ex)
            {
                email = "";
            }

            return email;
        }

        private static String AppendString(String main, String addOn)
        {
            if (main.Length > 0 && addOn.Length > 0)
            {
                main = main + ", " + addOn;
            }

            return main;
        }

        private static void WriteContactToFile(Contact thisContact, String emailAddress)
        {
            String fileName = "contact-" + emailAddress + ".txt";

            List<String> lines = new List<String>();

            foreach (var thisField in thisContact.GetLoadedPropertyDefinitions())
            {
                String fieldName = thisField.ToString();
                String value = ObjectToString(GetValue(thisContact, thisField));

                if (value.Equals("Microsoft.Exchange.WebServices.Data.EmailAddressDictionary"))
                {
                    String emailAddress1 = GetEmail(thisContact, EmailAddressKey.EmailAddress1);
                    String emailAddress2 = GetEmail(thisContact, EmailAddressKey.EmailAddress2);
                    String emailAddress3 = GetEmail(thisContact, EmailAddressKey.EmailAddress3);

                    value = emailAddress1;
                    value = AppendString(value, emailAddress2);
                    value = AppendString(value, emailAddress3);
                }
                else if (value.Equals("Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary"))
                {
                    value = "";
                }
                else if (value.Equals("Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary"))
                {
                    String phone1 = GetPhone(thisContact, PhoneNumberKey.PrimaryPhone);
                    String phone2 = GetPhone(thisContact, PhoneNumberKey.MobilePhone);
                    String phone3 = GetPhone(thisContact, PhoneNumberKey.BusinessPhone);
                    String phone4 = GetPhone(thisContact, PhoneNumberKey.CompanyMainPhone);
                    String phone5 = GetPhone(thisContact, PhoneNumberKey.OtherTelephone);

                    value = phone1;
                    value = AppendString(value, phone2);
                    value = AppendString(value, phone3);
                    value = AppendString(value, phone4);
                    value = AppendString(value, phone5);
                }

                lines.Add(fieldName + " : " + value);
            }

            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(fileName))
            {
                foreach (String line in lines)
                {
                    file.WriteLine(line);
                }
            }

            lines.Clear();
        }

        private static void GetContact(Contact thisContact, String emailAddress, MailboxType? mailboxType)
        {
            AddressClass newRecord = new AddressClass();

            Object displayName, alias, allowedResponseActions, archiveTag, assistantName, assistantPhone, attachments;
            Object birthday, businessAddressCity, businessAddressCountryOrRegion, businessAddressPostalCode;
            Object businessAddressState,
                businessAddressStreet,
                businessFax,
                businessHomePage,
                businessPhone,
                businessPhone2;
            Object carPhone, categories, children, companies, companyMainPhone, companyName;
            Object completeName, contactSource, conversationId, culture, dateTimeCreated, dateTimeReceived;
            Object dateTimeSent,
                department,
                directoryId,
                directReports,
                displayTo,
                emailAddress1,
                emailAddress2,
                emailAddress3;
            Object emailAddresses, entityExtractionResult, flag, generation, givenName, hasAttachments, hasPicture;
            Object homeAddressCity,
                homeAddressCountryOrRegion,
                homeAddressPostalCode,
                homeAddressState,
                homeAddressStreet,
                homeFax;
            Object homePhone, homePhone2, id, imAddress1, imAddress2, imAddress3, imAddresses;
            Object initials, instanceKey, isdn, jobTitle, lastModifiedName, lastModifiedTime;
            Object manager, managerMailbox, middleName, mileage, mobilePhone, nickName, notes, officeLocation;
            Object otherAddressCity,
                otherAddressCountryOrRegion,
                otherAddressPostalCode,
                otherAddressState,
                otherAddressStreet;
            Object otherFax, otherTelephone, pager, parentFolderId, phoneNumbers, phoneticFirstName;
            Object phoneticFullName, phoneticLastName, photo, physicalAddresses, policyTag, postalAddressIndex;
            Object primaryPhone, profession, radioPhone, spouseName, storeEntryId, surname;
            Object telex, ttyTddPhone, weddingAnniversary;

            if (String.IsNullOrEmpty(emailAddress))
            {
                // jmk found an issue where the email address was null. Couldn't figure out what this means.
                // The rest of the record was japaneese but the email address was indeed null.
                return;
            }

            if (!ContactExists(emailAddress))
            {
                newRecord.Address = emailAddress;
                newRecord.MailBoxType = mailboxType;
                AllContacts.Add(newRecord);
            }

            businessPhone = String.Empty;
            try
            {
                businessPhone = thisContact.PhoneNumbers[PhoneNumberKey.BusinessPhone].ToString();
            }
            catch (Exception ex)
            {
            }

            businessPhone2 = String.Empty;
            try
            {
                businessPhone2 = thisContact.PhoneNumbers[PhoneNumberKey.BusinessPhone2].ToString();
            }
            catch (Exception ex)
            {
            }

            emailAddress1 = String.Empty;
            try
            {
                emailAddress1 = thisContact.EmailAddresses[EmailAddressKey.EmailAddress1].Address;
            }
            catch (Exception ex)
            {
            }

            emailAddress2 = String.Empty;
            try
            {
                emailAddress2 = thisContact.EmailAddresses[EmailAddressKey.EmailAddress2].Address;
            }
            catch (Exception ex)
            {
            }

            emailAddress3 = String.Empty;
            try
            {
                emailAddress3 = thisContact.EmailAddresses[EmailAddressKey.EmailAddress3].Address;
            }
            catch (Exception ex)
            {
            }


            displayName = GetValue(thisContact, ContactSchema.DisplayName);
            alias = GetValue(thisContact, ContactSchema.Alias);
            //allowedResponseActions = GetValue(thisContact, ContactSchema.AllowedResponseActions);
            //archiveTag = GetValue(thisContact, ContactSchema.ArchiveTag);
            //assistantName = GetValue(thisContact, ContactSchema.AssistantName);
            //assistantPhone = GetValue(thisContact, ContactSchema.AssistantPhone);
            //attachments = GetValue(thisContact, ContactSchema.Attachments);
            //birthday = GetValue(thisContact, ContactSchema.Birthday);
            //businessAddressCity = GetValue(thisContact, ContactSchema.BusinessAddressCity);
            //businessAddressCountryOrRegion = GetValue(thisContact, ContactSchema.BusinessAddressCountryOrRegion);
            //businessAddressPostalCode = GetValue(thisContact, ContactSchema.BusinessAddressPostalCode);
            //businessAddressState = GetValue(thisContact, ContactSchema.BusinessAddressState);
            //businessAddressStreet = GetValue(thisContact, ContactSchema.BusinessAddressStreet);
            //businessFax = GetValue(thisContact, ContactSchema.BusinessFax);
            //businessHomePage = GetValue(thisContact, ContactSchema.BusinessHomePage);
            //businessPhone = GetValue(thisContact, ContactSchema.BusinessPhone);
            //businessPhone2 = GetValue(thisContact, ContactSchema.BusinessPhone2);
            //carPhone = GetValue(thisContact, ContactSchema.CarPhone);
            //categories = GetValue(thisContact, ContactSchema.Categories);
            //children = GetValue(thisContact, ContactSchema.Children);
            //companies = GetValue(thisContact, ContactSchema.Companies);
            //companyMainPhone = GetValue(thisContact, ContactSchema.CompanyMainPhone);
            companyName = GetValue(thisContact, ContactSchema.CompanyName);
            completeName = GetValue(thisContact, ContactSchema.CompleteName);
            contactSource = GetValue(thisContact, ContactSchema.ContactSource);
            //conversationId = GetValue(thisContact, ContactSchema.ConversationId);
            //culture = GetValue(thisContact, ContactSchema.Culture);
            //dateTimeCreated = GetValue(thisContact, ContactSchema.DateTimeCreated);
            //dateTimeReceived = GetValue(thisContact, ContactSchema.DateTimeReceived);
            //dateTimeSent = GetValue(thisContact, ContactSchema.DateTimeSent);
            department = GetValue(thisContact, ContactSchema.Department);
            directoryId = GetValue(thisContact, ContactSchema.DirectoryId);
            //directReports = GetValue(thisContact, ContactSchema.DirectReports);
            //displayName = GetValue(thisContact, ContactSchema.DisplayName);
            //displayTo = GetValue(thisContact, ContactSchema.DisplayTo);
            //emailAddress1 = GetValue(thisContact, ContactSchema.EmailAddress1);
            //emailAddress2 = GetValue(thisContact, ContactSchema.EmailAddress2);
            //emailAddress3 = GetValue(thisContact, ContactSchema.EmailAddress2);
            //emailAddresses = GetValue(thisContact, ContactSchema.EmailAddresses);
            //entityExtractionResult = GetValue(thisContact, ContactSchema.EntityExtractionResult);
            //flag = GetValue(thisContact, ContactSchema.Flag);
            //generation = GetValue(thisContact, ContactSchema.Generation);
            givenName = GetValue(thisContact, ContactSchema.GivenName);
            //hasAttachments = GetValue(thisContact, ContactSchema.HasAttachments);
            //hasPicture = GetValue(thisContact, ContactSchema.HasPicture);
            //homeAddressCity = GetValue(thisContact, ContactSchema.HomeAddressCity);
            //homeAddressCountryOrRegion = GetValue(thisContact, ContactSchema.HomeAddressCountryOrRegion);
            //homeAddressPostalCode = GetValue(thisContact, ContactSchema.HomeAddressPostalCode);
            //homeAddressState = GetValue(thisContact, ContactSchema.HomeAddressState);
            //homeAddressStreet = GetValue(thisContact, ContactSchema.HomeAddressStreet);
            //homeFax= GetValue(thisContact, ContactSchema.HomeFax);
            //homePhone = GetValue(thisContact, ContactSchema.HomePhone);
            //homePhone2 = GetValue(thisContact, ContactSchema.HomePhone2);
            //id = GetValue(thisContact, ContactSchema.Id);
            //imAddress1 = GetValue(thisContact, ContactSchema.ImAddress1);
            //imAddress2 = GetValue(thisContact, ContactSchema.ImAddress2);
            //imAddress3 = GetValue(thisContact, ContactSchema.ImAddress3);
            //imAddresses = GetValue(thisContact, ContactSchema.ImAddresses);
            //initials = GetValue(thisContact, ContactSchema.Initials);
            //instanceKey = GetValue(thisContact, ContactSchema.InstanceKey);
            //isdn = GetValue(thisContact, ContactSchema.Isdn);
            //jobTitle = GetValue(thisContact, ContactSchema.JobTitle);
            //lastModifiedName = GetValue(thisContact, ContactSchema.LastModifiedName);
            //lastModifiedTime = GetValue(thisContact, ContactSchema.LastModifiedTime);
            manager = GetValue(thisContact, ContactSchema.Manager);
            //managerMailbox = GetValue(thisContact, ContactSchema.ManagerMailbox);
            middleName = GetValue(thisContact, ContactSchema.MiddleName);
            //mileage = GetValue(thisContact, ContactSchema.Mileage);
            //mobilePhone = GetValue(thisContact, ContactSchema.MobilePhone);
            nickName = GetValue(thisContact, ContactSchema.NickName);
            notes = GetValue(thisContact, ContactSchema.Notes);
            //officeLocation = GetValue(thisContact, ContactSchema.OfficeLocation);
            //otherAddressCity = GetValue(thisContact, ContactSchema.OtherAddressCity);
            //otherAddressCountryOrRegion = GetValue(thisContact, ContactSchema.OtherAddressCountryOrRegion);
            //otherAddressPostalCode = GetValue(thisContact, ContactSchema.OtherAddressPostalCode);
            //otherAddressState = GetValue(thisContact, ContactSchema.OtherAddressState);
            //otherAddressStreet = GetValue(thisContact, ContactSchema.OtherAddressStreet);
            //otherFax = GetValue(thisContact, ContactSchema.OtherFax);
            //otherTelephone = GetValue(thisContact, ContactSchema.OtherTelephone);
            //pager = GetValue(thisContact, ContactSchema.Pager);
            //parentFolderId = GetValue(thisContact, ContactSchema.ParentFolderId);
            //phoneNumbers = GetValue(thisContact, ContactSchema.PhoneNumbers);
            //phoneticFirstName = GetValue(thisContact, ContactSchema.PhoneticFirstName);
            //phoneticFullName = GetValue(thisContact, ContactSchema.PhoneticFullName);
            //phoneticLastName = GetValue(thisContact, ContactSchema.PhoneticLastName);
            //photo = GetValue(thisContact, ContactSchema.Photo);
            //physicalAddresses = GetValue(thisContact, ContactSchema.PhysicalAddresses);
            //policyTag = GetValue(thisContact, ContactSchema.PolicyTag);
            //postalAddressIndex = GetValue(thisContact, ContactSchema.PostalAddressIndex);
            //primaryPhone = GetValue(thisContact, ContactSchema.PrimaryPhone);
            //profession = GetValue(thisContact, ContactSchema.Profession);
            //radioPhone = GetValue(thisContact, ContactSchema.RadioPhone);
            //spouseName = GetValue(thisContact, ContactSchema.SpouseName);
            //storeEntryId = GetValue(thisContact, ContactSchema.StoreEntryId);
            surname = GetValue(thisContact, ContactSchema.Surname);
            //telex = GetValue(thisContact, ContactSchema.Telex);
            //ttyTddPhone = GetValue(thisContact, ContactSchema.TtyTddPhone);
            //weddingAnniversary = GetValue(thisContact, ContactSchema.WeddingAnniversary);

            newRecord.DisplayName = ObjectToString(displayName);
            newRecord.Alias = ObjectToString(alias);
            //newRecord.AllowedResponseActions = ObjectToString(allowedResponseActions);
            //newRecord.ArchiveTag = ObjectToString(archiveTag);
            //newRecord.AssistantName = ObjectToString(assistantName);
            //newRecord.AssistantPhone = ObjectToString(assistantPhone);
            //newRecord.Attachments = ObjectToString(attachments);
            //newRecord.Birthday = ObjectToString(birthday);
            //newRecord.BusinessAddressCity = ObjectToString(businessAddressCity);
            //newRecord.BusinessAddressCountryOrRegion = ObjectToString(businessAddressCountryOrRegion);
            //newRecord.BusinessAddressPostalCode = ObjectToString(businessAddressPostalCode);
            //newRecord.BusinessAddressState = ObjectToString(businessAddressState);
            //newRecord.BusinessAddressStreet = ObjectToString(businessAddressStreet);
            //newRecord.BusinessFax = ObjectToString(businessFax);
            //newRecord.BusinessHomePage = ObjectToString(businessHomePage);
            newRecord.BusinessPhone = ObjectToString(businessPhone);
            newRecord.BusinessPhone2 = ObjectToString(businessPhone2);
            //newRecord.CarPhone = ObjectToString(carPhone);
            //newRecord.Categories = ObjectToString(categories);
            //newRecord.Children = ObjectToString(children);
            //newRecord.Companies = ObjectToString(companies);
            //newRecord.CompanyMainPhone = ObjectToString(companyMainPhone);
            newRecord.CompanyName = ObjectToString(companyName);
            newRecord.CompleteName = ObjectToString(completeName);
            newRecord.ContactSource = ObjectToString(contactSource);
            //newRecord.ConversationId = ObjectToString(conversationId);
            //newRecord.Culture = ObjectToString(culture);
            //newRecord.DateTimeCreated = ObjectToString(dateTimeCreated);
            //newRecord.DateTimeReceived = ObjectToString(dateTimeReceived);
            //newRecord.DateTimeSent = ObjectToString(dateTimeSent);
            newRecord.Department = ObjectToString(department);
            newRecord.DirectoryId = ObjectToString(directoryId);
            //newRecord.DirectReports = ObjectToString(directReports);
            newRecord.DisplayName = ObjectToString(displayName);
            //newRecord.DisplayTo = ObjectToString(displayTo);
            newRecord.EmailAddress1 = ObjectToString(emailAddress1);
            newRecord.EmailAddress2 = ObjectToString(emailAddress2);
            newRecord.EmailAddress3 = ObjectToString(emailAddress3);
            //newRecord.EmailAddresses = ObjectToString(emailAddresses);
            //newRecord.EntityExtractionResult = ObjectToString(entityExtractionResult);
            //newRecord.Flag = ObjectToString(flag);
            //newRecord.Generation = ObjectToString(generation);
            newRecord.GivenName = ObjectToString(givenName);
            //newRecord.HasAttachments = ObjectToString(hasAttachments);
            //newRecord.HasPicture = ObjectToString(hasPicture);
            //newRecord.HomeAddressCity = ObjectToString(homeAddressCity);
            //newRecord.HomeAddressCountryOrRegion = ObjectToString(homeAddressCountryOrRegion);
            //newRecord.HomeAddressPostalCode = ObjectToString(homeAddressPostalCode);
            //newRecord.HomeAddressState = ObjectToString(homeAddressState);
            //newRecord.HomeAddressStreet = ObjectToString(homeAddressStreet);
            //newRecord.HomeFax = ObjectToString(homeFax);
            //newRecord.HomePhone = ObjectToString(homePhone);
            //newRecord.HomePhone2 = ObjectToString(homePhone2);
            //newRecord.Id = ObjectToString(id);
            //newRecord.ImAddress1 = ObjectToString(imAddress1);
            //newRecord.ImAddress2 = ObjectToString(imAddress2);
            //newRecord.ImAddress3 = ObjectToString(imAddress3);
            //newRecord.ImAddresses = ObjectToString(imAddresses);
            //newRecord.Initials = ObjectToString(initials);
            //newRecord.InstanceKey = ObjectToString(instanceKey);
            //newRecord.Isdn = ObjectToString(isdn);
            //newRecord.JobTitle = ObjectToString(jobTitle);
            //newRecord.LastModifiedName = ObjectToString(lastModifiedName);
            //newRecord.LastModifiedTime = ObjectToString(lastModifiedTime);
            newRecord.Manager = ObjectToString(manager);
            //newRecord.ManagerMailbox = ObjectToString(managerMailbox);
            newRecord.MiddleName = ObjectToString(middleName);
            //newRecord.Mileage = ObjectToString(mileage);
            //newRecord.MobilePhone = ObjectToString(mobilePhone);
            newRecord.NickName = ObjectToString(nickName);
            newRecord.Notes = ObjectToString(notes);
            //newRecord.OfficeLocation = ObjectToString(officeLocation);
            //newRecord.OtherAddressCity = ObjectToString(otherAddressCity);
            //newRecord.OtherAddressCountryOrRegion = ObjectToString(otherAddressCountryOrRegion);
            //newRecord.OtherAddressPostalCode = ObjectToString(otherAddressPostalCode);
            //newRecord.OtherAddressState = ObjectToString(otherAddressState);
            //newRecord.OtherAddressStreet = ObjectToString(otherAddressStreet);
            //newRecord.OtherFax = ObjectToString(otherFax);
            //newRecord.OtherTelephone = ObjectToString(otherTelephone);
            //newRecord.Pager = ObjectToString(pager);
            //newRecord.ParentFolderId = ObjectToString(parentFolderId);
            //newRecord.PhoneNumbers = ObjectToString(phoneNumbers);
            //newRecord.PhoneticFirstName = ObjectToString(phoneticFirstName);
            //newRecord.PhoneticFullName = ObjectToString(phoneticFullName);
            //newRecord.PhoneticLastName = ObjectToString(phoneticLastName);
            //newRecord.Photo = ObjectToString(photo);
            //newRecord.PhysicalAddresses = ObjectToString(physicalAddresses);
            //newRecord.PolicyTag = ObjectToString(policyTag);
            //newRecord.PostalAddressIndex = ObjectToString(postalAddressIndex);
            //newRecord.PrimaryPhone = ObjectToString(primaryPhone);
            //newRecord.Profession = ObjectToString(profession);
            //newRecord.RadioPhone = ObjectToString(radioPhone);
            //newRecord.SpouseName = ObjectToString(spouseName);
            //newRecord.StoreEntryId = ObjectToString(storeEntryId);
            newRecord.Surname = ObjectToString(surname);
            //newRecord.Telex = ObjectToString(telex);
            //newRecord.TtyTddPhone = ObjectToString(ttyTddPhone);
            //newRecord.WeddingAnniversary = ObjectToString(weddingAnniversary);
        }

        private static Boolean GroupExists(String groupName)
        {
            foreach (String thisOne in AllFoundGroupNames)
            {
                if (thisOne.Equals(groupName))
                {
                    return true;
                }
            }

            return false;
        }

        private static void GetMembersForGroup(ExchangeService exchangeService, string groupName)
        {
            ExpandGroupResults myGroupMembers;

            if (GroupExists(groupName).Equals(true))
            {
                return;
            }
            else
            {
                AllFoundGroupNames.Add(groupName);
            }

            try
            {
                myGroupMembers = exchangeService.ExpandGroup(groupName);
            }
            catch (Exception ex)
            {
                return;
            }

            if (myGroupMembers.Count().Equals(0))
            {
                return;
            }

            String fileName = groupName.ToLower() + ".groupmembers";
            Console.WriteLine();
            Console.WriteLine("    Group:" + groupName + " Members:" + myGroupMembers.Count());
            fileName = fileName.Replace("/", "_");
            fileName = fileName.Replace("*", "_");
            fileName = fileName.Replace(":", "_");
            fileName = fileName.Replace("<", "_");
            fileName = fileName.Replace(">", "_");
            fileName = fileName.Replace(@"\", "_");
            fileName = fileName.Replace("|", "_");
            fileName = fileName.Replace("?", "_");
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@fileName))
            {
                if (myGroupMembers.IncludesAllMembers.Equals(false))
                {
                    Console.WriteLine("    -truncated list");
                }

                foreach (EmailAddress address in myGroupMembers.Members)
                {
                    file.WriteLine(address.Address);
                }
            }
        }

        private static void DumpGal(ExchangeService exchangeService,
            string search,
            Boolean fast,
            Boolean createContactsFiles,
            Boolean createGroupFiles)
        {
            List<FolderId> folders = new List<FolderId>() {new FolderId(WellKnownFolderName.Contacts)};
            String name = "", emailAddress = "", mailboxType = "";
            NameResolutionCollection coll = null;

            try
            {
                coll = exchangeService.ResolveName(search, folders, ResolveNameSearchLocation.DirectoryOnly, true);
            }
            catch (System.ArgumentException ex)
            {
                Console.WriteLine("{SEARCH=" + search +
                                  "} There is a junk record in this set. Going to try and expand out.");
                foreach (string thisOne in StartArray)
                {
                    DumpGal(exchangeService, search + thisOne, fast, createContactsFiles, createGroupFiles);
                }
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine(ex.ToString());
                return;
            }
            catch (ServiceResponseException ex)
            {
                Console.WriteLine("Error:");
                Console.WriteLine("The connection has timed out.");
                Console.WriteLine(ex.ToString());
                Console.WriteLine("Going to wait 15 seconds and try again");
                System.Threading.Thread.Sleep(15000);
                DumpGal(exchangeService, search, fast, createContactsFiles, createGroupFiles);
                return;
            }

            DumpOutput();

            if (coll != null)
            {
                if (coll.Count > 0)
                {
                    if (coll.IncludesAllResolutions.Equals(false) && fast.Equals(false))
                    {
                        Console.WriteLine();
                        Console.WriteLine("(" + search + ") (" + coll.Count +
                                          ") This does not include all items. Expanding");

                        foreach (string thisOne in StartArray)
                        {
                            DumpGal(exchangeService, search + thisOne, fast, createContactsFiles, createGroupFiles);
                        }
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine(search + " : " + coll.Count);
                        foreach (NameResolution nameRes in coll)
                        {
                            Contact thisContact = nameRes.Contact;
                            EmailAddress thisMailbox = nameRes.Mailbox;

                            if (thisMailbox != null)
                            {
                                name = thisMailbox.Name;
                                emailAddress = thisMailbox.Address;
                                mailboxType = thisMailbox.MailboxType.ToString();
                            }

                            if (thisContact != null)
                            {
                                switch (thisMailbox.MailboxType)
                                {
                                    case MailboxType.Mailbox:
                                    case MailboxType.Contact:
                                        GetContact(thisContact, emailAddress, thisMailbox.MailboxType);
                                        if (createContactsFiles.Equals(true))
                                        {
                                            WriteContactToFile(thisContact, emailAddress);
                                        }

                                        break;
                                    case MailboxType.PublicGroup:
                                    case MailboxType.ContactGroup:
                                        if (fast.Equals(false))
                                        {
                                            if (createGroupFiles.Equals(true))
                                            {
                                                GetMembersForGroup(exchangeService, name);
                                            }
                                        }

                                        break;
                                    case MailboxType.OneOff:
                                        Console.WriteLine();
                                        Console.WriteLine("OneOff - (" + name + ")");
                                        break;
                                    case MailboxType.PublicFolder:
                                        Console.WriteLine();
                                        Console.WriteLine("PublicFolder - (" + name + ")");
                                        break;
                                    case MailboxType.Unknown:
                                        Console.WriteLine();
                                        Console.WriteLine("Unknown - (" + name + ")?");
                                        break;
                                }
                            }
                        }
                    }
                }
            }
        }

        private static void DisplayStandardErrorInformation(String location)
        {
            Console.WriteLine("There was an error trying to find the Autodiscover service (" + location + ")");
            Console.WriteLine("This could be due to:");
            Console.WriteLine("   1) There isn't an Exchange server for this domain");
            Console.WriteLine("   2) The username/password/e-mail isn't correct");
            Console.WriteLine("   3) The Exchange Server isn't >= v2007");
            Console.WriteLine("   4) The Autodiscover configuration is strange");
            Console.WriteLine("   5) Windows doesn't have everything it needs to run this application");
            Console.WriteLine();
            Console.WriteLine("Error: ");
        }

        private static System.Uri FindAutodiscoverUrl(String emailAddress, String login, String password)
        {
            String shortMessage = String.Empty;
            ExchangeService service = null;

            try
            {
                service = new ExchangeService(ExchangeVersion.Exchange2010);
                service.Credentials = new WebCredentials(login, password);
                service.AutodiscoverUrl(emailAddress, RedirectionCallback);
                Console.WriteLine("     EWS Endpoint: {0}", service.Url);
            }
            catch (AutodiscoverRemoteException ex)
            {
                shortMessage = CleanError(ex.ToString());
                DisplayStandardErrorInformation("AutodiscoverRemoteException");
                Console.WriteLine(shortMessage);
                Environment.Exit(0);
            }
            catch (AutodiscoverLocalException ex)
            {
                shortMessage = CleanError(ex.ToString());
                DisplayStandardErrorInformation("AutodiscoverLocalException");
                Console.WriteLine(shortMessage);
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                shortMessage = CleanError(ex.ToString());
                DisplayStandardErrorInformation("General Exception");
                Console.WriteLine(shortMessage);
                Environment.Exit(0);
            }

            return service.Url;
        }

        private static String CleanError(String error)
        {
            if (error.Length.Equals(0))
            {
                return "";
            }

            Int32 firstLocation, lastLocation, length;
            String cleanString = String.Empty;

            firstLocation = error.IndexOf(": ");
            if (firstLocation > -1)
            {
                firstLocation = firstLocation + 2;
            }

            lastLocation = error.IndexOf(" at ");
            length = lastLocation - firstLocation;

            if (firstLocation > -1 && lastLocation > -1 && length > 0)
            {
                cleanString = error.Substring(firstLocation, length);
            }
            else
            {
                cleanString = error;
            }

            return cleanString;
        }

        private static ExchangeService TestVersion(ExchangeVersion versionToTest, String login, String password,
            System.Uri exchangeInformation)
        {
            ExchangeService _service = null;

            try
            {
                _service = new ExchangeService(versionToTest);
                _service.Credentials = new WebCredentials(login, password);
                _service.Url = exchangeInformation;


                ServicePointManager.ServerCertificateValidationCallback =
                    delegate(object sender1,
                        System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                        System.Security.Cryptography.X509Certificates.X509Chain chain,
                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
                    {
                        return true;
                    };

                List<FolderId> folders = new List<FolderId>() {new FolderId(WellKnownFolderName.Contacts)};
                NameResolutionCollection coll =
                    _service.ResolveName("xx", folders, ResolveNameSearchLocation.DirectoryOnly, false);
            }
            catch (ServiceVersionException ex)
            {
                _service = null;
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Error connecting");
                String shortMessage = CleanError(ex.ToString());

                if (ex.Message.IndexOf("Unauthorized") > -1)
                {
                    Console.WriteLine("Server returning 'Unauthorized'.  Bad password maybe?");
                    Console.WriteLine("Error: ");
                    Console.WriteLine(shortMessage);
                    _service = null;
                    Environment.Exit(0);
                }
                else
                {
                    Console.WriteLine("Error: ");
                    Console.WriteLine(shortMessage);
                    _service = null;
                    Environment.Exit(0);
                }
            }
            catch (ServiceResponseException ex)
            {
                Console.WriteLine("Error connecting");
                String shortMessage = CleanError(ex.ToString());
                Console.WriteLine("Error: " + shortMessage);
                _service = null;
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error connecting");
                Console.WriteLine(ex.ToString());
                _service = null;
            }

            return _service;
        }

        private static bool RedirectionCallback(string url)
        {
            // Return true if the URL is an HTTPS URL.
            return url.ToLower().StartsWith("https://");
        }

        private static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("dru1d's dotnet core implementation of gal-dump");
            Console.WriteLine("Usage: galdump [OPTIONS]+");
            Console.WriteLine("Example: galdump -e psyonik@foofus.net -l david -p Password1");
            Console.WriteLine();
            Console.WriteLine("galdump is an application that extracts the GAL from");
            Console.WriteLine("an Exchange server using EWS (derived from the e-mail address).");
            Console.WriteLine("Exchange server versions require 2007 or greater.");
            Console.WriteLine();
            Console.WriteLine("Notes:");
            Console.WriteLine("   Office365 EWS Location: https://outlook.office365.com/EWS/Exchange.asmx");
            Console.WriteLine("    - Login name would be the email address");
            Console.WriteLine("   If you need the domain to login, use DOMAIN\\USERNAME");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        private static void DisplayEwsVersion(ExchangeService exchangeService)
        {
            String exchangeServerVersionFromServer;

            ExchangeVersion dwdw = exchangeService.RequestedServerVersion;
            exchangeServerVersionFromServer = dwdw.ToString();

            ExchangeServerInfo jojo = exchangeService.ServerInfo;
            exchangeServerVersionFromServer = exchangeServerVersionFromServer + " (" + jojo.MajorVersion + "." +
                                              jojo.MinorVersion + "." + jojo.MajorBuildNumber + "." +
                                              jojo.MinorBuildNumber + ")";

            Console.WriteLine(exchangeServerVersionFromServer);
        }

        private static void DumpOutput()
        {
            String fileName = "galdump.all";

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@fileName))
            {
                file.WriteLine(
                    "Email|Display Name|Given Name|Middle Name|Surname|Job Title|Manager|Nick Name|Notes|Alias|Business Phone|Business Phone 2|Email Address 1|Email Address 2|Email Address 3");
                foreach (AddressClass thisAddress in AllContacts)
                {
                    file.WriteLine(thisAddress.Address + "|" + thisAddress.DisplayName + "|" + thisAddress.GivenName +
                                   "|" + thisAddress.MiddleName + "|" + thisAddress.Surname + "|" +
                                   thisAddress.JobTitle + "|" + thisAddress.Manager + "|" + thisAddress.NickName + "|" +
                                   thisAddress.Notes + "|" + thisAddress.Alias + "|" + thisAddress.BusinessPhone + "|" +
                                   thisAddress.BusinessPhone2 + "|" + thisAddress.EmailAddress1 + "|" +
                                   thisAddress.EmailAddress2 + "|" + thisAddress.EmailAddress3);
                }
            }
        }

        static void Main(string[] args)
        {
            bool showHelp = false;
            List<string> names = new List<string>();
            String emailAddress = "";
            String password = "";
            String login = "";
            String ewsUrl = "";
            ExchangeService service = null;
            Boolean found = false;
            Boolean fast = false;
            Boolean createContactsFiles = false;
            Boolean createGroupFiles = false;
            System.Uri exchangeServerInformation;

            OptionSet p = new OptionSet()
            {
                {"e|email=", "A known (valid) e-mail address (if ews url is not set)", v => emailAddress = v},
                {"l|login=", "Login account to the exchange server", v => login = v},
                {"p|password=", "Password for the login account", v => password = v},
                {"x|ewsurl=", "Full URL to the EWS service (https://a.b.c/ews/Exchange.asmx)", v => ewsUrl = v},
                {"f|fast=", "Don't worry about getting all records", v => fast = true},
                {"c|contacts=", "Create contact files", v => createContactsFiles = true},
                {"g|groups=", "Create group files", v => createGroupFiles = true},
                {"h|help", "show this message and exit", v => showHelp = v != null},
            };

            try
            {
                p.Parse(args);
            }
            catch (OptionException e)
            {
                Console.Write("galdump: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `galdump --help' for more information.");
                return;
            }


            if (showHelp.Equals(false))
            {
                if (emailAddress.Equals(""))
                {
                    // Email address only needed if the ews path is not set (for lookup purposes)
                    if (ewsUrl.Equals(""))
                    {
                        Console.WriteLine("Missing e-mail address");
                        showHelp = true;
                    }
                }

                if (login.Equals(""))
                {
                    Console.WriteLine("Missing login account name");
                    showHelp = true;
                }

                if (password.Equals(""))
                {
                    Console.WriteLine("Missing password");
                    showHelp = true;
                }
            }

            if (showHelp)
            {
                ShowHelp(p);
                return;
            }

            ServicePointManager.ServerCertificateValidationCallback = delegate(Object obj, X509Certificate certificate,
                X509Chain chain, SslPolicyErrors errors)
            {
                return true;
                
                if (errors == System.Net.Security.SslPolicyErrors.None)
                {
                    return true;
                }

                if ((errors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
                {
                    if (chain != null && chain.ChainStatus != null)
                    {
                        foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain
                            .ChainStatus)
                        {
                            if ((certificate.Subject == certificate.Issuer) &&
                                (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags
                                     .UntrustedRoot))
                            {
                                continue;
                            }
                            else
                            {
                                if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags
                                        .NoError)
                                {
                                    return false;
                                }
                            }
                        }
                    }

                    return true;
                }
                else
                {
                    return false;
                }
            };

            if (ewsUrl.Equals(String.Empty))
            {
                Console.WriteLine("Looking for Exchange Server EWS location for {0}", emailAddress);
                exchangeServerInformation = FindAutodiscoverUrl(emailAddress, login, password);
            }
            else
            {
                exchangeServerInformation = new System.Uri(ewsUrl);
            }

            Console.WriteLine("Testing for Exchange Version");

            if (found.Equals(false))
            {
                service = TestVersion(ExchangeVersion.Exchange2013, login, password, exchangeServerInformation);
                if (service != null)
                {
                    Console.WriteLine("     Version Exchange 2013");
                    found = true;
                }
                else
                {
                    Console.Write(".");
                }
            }

            if (found.Equals(false))
            {
                service = TestVersion(ExchangeVersion.Exchange2010_SP2, login, password, exchangeServerInformation);
                if (service != null)
                {
                    Console.WriteLine("     Version Exchange 2010 SP2");
                    found = true;
                }
                else
                {
                    Console.Write(".");
                }
            }

            if (found.Equals(false))
            {
                service = TestVersion(ExchangeVersion.Exchange2010_SP1, login, password, exchangeServerInformation);
                if (service != null)
                {
                    Console.WriteLine("     Version Exchange 2010 SP1");
                    found = true;
                }
                else
                {
                    Console.Write(".");
                }
            }

            if (found.Equals(false))
            {
                service = TestVersion(ExchangeVersion.Exchange2010, login, password, exchangeServerInformation);
                if (service != null)
                {
                    Console.WriteLine("     Version Exchange 2010");
                    found = true;
                }
                else
                {
                    Console.Write(".");
                }
            }

            if (found.Equals(false))
            {
                service = TestVersion(ExchangeVersion.Exchange2007_SP1, login, password, exchangeServerInformation);
                if (service != null)
                {
                    Console.WriteLine("     Version Exchange 2007 SP1");
                    found = true;
                }
                else
                {
                    Console.Write(".");
                }
            }


            if (found.Equals(true))
            {
                foreach (string thisOne in StartArray)
                {
                    DumpGal(service, thisOne, fast, createContactsFiles, createGroupFiles);
                }

                DumpOutput();
            }
        }
    }
}