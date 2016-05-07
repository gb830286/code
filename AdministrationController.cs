using BECOR_Website.BECOR_EntityDataModel;
using BECOR_Website.Email;
using BECOR_Website.Enums;
using BECOR_Website.HelperClasses;
using BECOR_Website.Mappers;
using BECOR_Website.Mappers.Administration;
using BECOR_Website.Models;
using BECOR_Website.Models.Administration;
using BECOR_Website.Resources;
using MailChimp.Campaigns;
using MailChimp.Helper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Mvc;
using MailChimp.Errors;
using System.Data.SqlClient;
using System.Data;

namespace BECOR_Website.Controllers.Administration
{
    [Authorize(Roles= "Administrator")]
    public class AdministrationController : MasterController
    {
        public ActionResult Index()
        {
            return RedirectToAction("AccountManagement");
        } 
       
        #region Account
        public ActionResult AccountManagement(ManageAccountFilterTypes? filterType)
        {
            List<UserProfileViewModel> viewModel;
            ViewBag.filterType = filterType;

            if (filterType != null)
            {
                viewModel = ef.UserProfiles
                .Where(up => (filterType == ManageAccountFilterTypes.Active ? up.IsActive : !up.IsActive
                    && up.MembershipTypeID == (short)MembershipTypes.NonMember
                    || up.MembershipTypeID == (short)MembershipTypes.Individual
                    || up.MembershipTypeID == (short)MembershipTypes.Student
                    ))
                .ToList()
                .Select(up => UserProfileMapper.ToVM(up, CurrentCulture))
                .ToList();
            }
            else
            {
                viewModel = ef.UserProfiles
                 .Where(up => up.MembershipTypeID == (short)MembershipTypes.NonMember
                    || up.MembershipTypeID == (short)MembershipTypes.Individual
                    || up.MembershipTypeID == (short)MembershipTypes.Student)
                .ToList()
                .Select(up => UserProfileMapper.ToVM(up, CurrentCulture))
                .ToList();
            }

            foreach (var item in viewModel)
            {
                if (item.ParentUserGuid != null)
                {
                    item.ParentUserEmail = ef.UserProfiles.Find(item.ParentUserGuid).EmailAddress;
                }
            }
            return View(viewModel);
        }

        public ActionResult AccountManagementCorporate()
        {
            List<UserProfileViewModel> viewModel = ef.UserProfiles
                .Where(up => up.MembershipTypeID == (short)MembershipTypes.Corporate
                    || up.MembershipTypeID == (short)MembershipTypes.CorporatePremium
                    || up.MembershipTypeID == (short)MembershipTypes.CorporatePremiumPlus)
                .ToList()
                .Select(up => UserProfileMapper.ToVM(up, CurrentCulture))
                .ToList();

            return View(viewModel);
        }

        public ActionResult AccountManagementBecor()
        {
            List<UserProfileViewModel> viewModel = ef.UserProfiles
                .Where(up => up.MembershipTypeID == (short)MembershipTypes.BECOR
                             || up.MembershipTypeID == (short)MembershipTypes.Administrator)
                .ToList()
                .Select(up => UserProfileMapper.ToVM(up, CurrentCulture))
                .ToList();

            return View(viewModel);
        }

        #region User Membership
        [Authorize(Roles = "Administrator")]
        public ActionResult EditUserMembership(string id, string returnUrl)
        {
            string deEmail = Encrypter.DecryptToString(id);
            if (deEmail == BECORConstants.ADMIN_EMAIL_ADDRESS)
            {
                return RedirectToAction("Index", "Error");
            }

            UserProfile userProfile = ef.UserProfiles.Where(up => up.EmailAddress == deEmail)
                .FirstOrDefault();

            ChangeMembershipViewModel viewModel = new ChangeMembershipViewModel
            { 
                DisplayEmail = userProfile.EmailAddress,
                Email = Encrypter.Encrypt(userProfile.EmailAddress), 
                MembershipTypeID = Encrypter.Encrypt(userProfile.MembershipTypeID),
                ReturnUrl = returnUrl,
            };

            ViewBag.MembershipList = Utilities.ConvertToSelectList(GetListOfMembershipTypes());
            return View(viewModel);
        }

        [Authorize(Roles = "Administrator")]
        [HttpPost]
        public ActionResult EditUserMembership(ChangeMembershipViewModel viewModel)
        {
            if(ModelState.IsValid)
            {
                string deEmail = Encrypter.DecryptToString(viewModel.Email);
                if (deEmail == BECORConstants.ADMIN_EMAIL_ADDRESS)
                {
                    return RedirectToAction("Index", "Error");
                }

                UserProfile userProfile = ef.UserProfiles
                    .Where(up => up.EmailAddress == deEmail)
                    .FirstOrDefault();
                UserProfileViewModel userProfileVM = UserProfileMapper.ToVM(userProfile, CurrentCulture);

                userProfileVM.MembershipTypeID = (short)Encrypter.DecryptToInt(viewModel.MembershipTypeID);
                if (userProfileVM.MembershipTypeID == (short)MembershipTypes.Administrator
                    || userProfileVM.MembershipTypeID == (short)MembershipTypes.Administrator
                    || userProfileVM.MembershipTypeID == (short)MembershipTypes.NonMember)
                {
                    userProfileVM.MembershipExpiration = null; 
                }
                else
                {
                    //TODO: FUTURE DEVELOPER - Unknown if Student Accounts expire at this time. However we have set them to expire at the end of fiscal year with the rest.
                    userProfileVM.MembershipExpiration = BECORConstants.MEMBERSHIP_EXPIRATION;
                }

                try
                {
                    ef.Entry(userProfile).CurrentValues.SetValues(UserProfileMapper.ToEF(userProfileVM));
                    ef.SaveChanges();

                    object[] args = new object[] { userProfile.EmailAddress, userProfileVM.MembershipTypeName };
                    StoreSuccessMessage(string.Format(SuccessMessages.UserUpdated, args));
                    return RedirectToLocal(viewModel.ReturnUrl);
                }
                catch (Exception)
                {
                    return RedirectToAction("Index", "Error");
                }
            }

            StoreModelState();
            return RedirectToAction("EditUserMembership", new { email = viewModel.Email });
        }
        #endregion

        #region Deactivate/Reactivate
        public ActionResult DeactivateAccount(string id, string returnUrl)
        {
            string deEmail = Encrypter.DecryptToString(id);
            if(deEmail == BECORConstants.ADMIN_EMAIL_ADDRESS)
            {
                return RedirectToAction("Index", "Error");
            }

            UserProfile userProfile = ef.UserProfiles.Where(up => up.EmailAddress == deEmail)
                .FirstOrDefault();
            UserProfileViewModel viewModel = UserProfileMapper.ToVM(userProfile, CurrentCulture);
            if (userProfile != null && userProfile.IsActive == true)
            {
                try
                {
                    viewModel.IsActive = false;
                    viewModel.IsSubscribedToMailingList = false;
                    ef.Entry(userProfile).CurrentValues.SetValues(UserProfileMapper.ToEF(viewModel));
                    ef.SaveChanges();
                    

                    StoreSuccessMessage(string.Format(SuccessMessages.AccountDeactivated, userProfile.EmailAddress));
                    return RedirectToLocal(returnUrl);
                }
                catch (Exception)
                {
                    return RedirectToAction("Index", "Error");
                }
            }

            StoreModelState("", ModelStateErrors.GeneralError);
            return RedirectToLocal(returnUrl);
        }

        public ActionResult ReactivateAccount(string id, string returnUrl)
        {
            string deEmail = Encrypter.DecryptToString(id);
            if (deEmail == BECORConstants.ADMIN_EMAIL_ADDRESS)
            {
                return RedirectToAction("Index", "Error");
            }

            UserProfile userProfile = ef.UserProfiles.Where(up => up.EmailAddress == deEmail)
                .FirstOrDefault();
            UserProfileViewModel viewModel = UserProfileMapper.ToVM(userProfile, CurrentCulture);
            if (userProfile != null && userProfile.IsActive == false)
            {
                try
                {
                    viewModel.IsActive = true;
                    ef.Entry(userProfile).CurrentValues.SetValues(UserProfileMapper.ToEF(viewModel));
                    ef.SaveChanges();

                    StoreSuccessMessage(string.Format(SuccessMessages.AccountReactivated, userProfile.EmailAddress));
                    return RedirectToLocal(returnUrl);
                }
                catch (Exception)
                {
                    return RedirectToAction("Index", "Error");
                }

            }

            StoreModelState("", ModelStateErrors.GeneralError);
            return RedirectToLocal(returnUrl);
        }
        #endregion        

        #region Migrate Account Owner
        public ActionResult MigrateAccountOwner(string id, string returnUrl)
        {
            string deEmail = Encrypter.DecryptToString(id);
            if (deEmail == BECORConstants.ADMIN_EMAIL_ADDRESS)
            {
                return RedirectToAction("Index", "Error");
            }

            UserProfile currentProfile = ef.UserProfiles.Where(cup => cup.EmailAddress == deEmail)
                .FirstOrDefault();

            UserMigrationViewModel viewModel = new UserMigrationViewModel();
            if(currentProfile != null)
            {
                viewModel.OldEmailAddressDisplay = currentProfile.EmailAddress;
                viewModel.OldEmailAddress = Encrypter.Encrypt(currentProfile.EmailAddress);
                viewModel.NewEmailAddress = null;
                viewModel.FirstName = currentProfile.FirstName;
                viewModel.LastName = currentProfile.LastName;
                viewModel.ReturnUrl = returnUrl;
            }
            return View(viewModel);
        }

        [HttpPost]
        public ActionResult MigrateAccountOwner(UserMigrationViewModel viewModel)
        {            
            if(ModelState.IsValid)
            {
                string deEmail = Encrypter.DecryptToString(viewModel.OldEmailAddress);
                if (deEmail == BECORConstants.ADMIN_EMAIL_ADDRESS)
                {
                    return RedirectToAction("Index", "Error");
                }
                
                // Check old email
                AspNetUser aspProfile = ef.AspNetUsers.Where(ap => ap.UserName == deEmail)
                    .FirstOrDefault();
                UserProfile currentProfile = ef.UserProfiles.Where(cup => cup.EmailAddress == deEmail)
                    .FirstOrDefault();

                // Check if new email exists
                int doesNewEmailExists = ef.UserProfiles.Where(cup => cup.EmailAddress == viewModel.NewEmailAddress)
                    .Count();

                if (doesNewEmailExists > 0)
                {
                    StoreModelState("", string.Format(ModelStateErrors.EmailExistsCrop, viewModel.NewEmailAddress));
                    return RedirectToLocal(viewModel.ReturnUrl);
                }

                if(aspProfile != null && currentProfile != null)
                {
                    object[] args = new object[] { currentProfile.EmailAddress, viewModel.NewEmailAddress };
                    aspProfile.UserName = viewModel.NewEmailAddress;
                    aspProfile.Email = viewModel.NewEmailAddress;
                    aspProfile.EmailConfirmed = false; 

                    currentProfile.EmailAddress = viewModel.NewEmailAddress;
                    currentProfile.FirstName = viewModel.FirstName;
                    currentProfile.LastName = viewModel.LastName;

                    try
                    {
                        ef.SaveChanges();

                        string password = System.Web.Security.Membership.GeneratePassword(12, 3);
                        string code = UserManager.GenerateEmailConfirmationTokenAsync(currentProfile.UserGuid).ToString();
                        var callbackUrl = Url.Action("ConfirmEmail", "Account", new { userId = currentProfile.UserGuid, code }, protocol: Request.Url.Scheme);

                        TaskFactory factory = new TaskFactory();
                        var appEmail = new ApplicationEmail();
                        factory.StartNew(() => appEmail.SendMail("Register Account", callbackUrl, currentProfile.EmailAddress, 
                            ApplicationEmailEnum.MigrateAccountOwner, password));

                        StoreSuccessMessage(string.Format(SuccessMessages.AccountMigrate, args));
                        return RedirectToLocal(viewModel.ReturnUrl);
                    }
                    catch (Exception)
                    {
                        return RedirectToAction("Index", "Error");
                    }
                }
                StoreModelState("", ModelStateErrors.GeneralError);
                return RedirectToLocal(viewModel.ReturnUrl);
            }  
            StoreModelState();
            return RedirectToAction("MigrateAccountOwner", new { id = viewModel.OldEmailAddress });
        }
        #endregion

        #endregion

        #region Membership
        public ActionResult ManageMembershipTypes()
        {
            List<MembershipTypeViewModel> viewModel = ef.MembershipTypes
                .ToList()
                .Select(mt => MembershipMapper.ToVM(mt, CurrentCulture))
                .ToList();

            return View(viewModel);
        }

        public ActionResult EditMembershipType(string id)
        {
            MembershipTypeViewModel viewModel = GetViewModel() != null
                ? GetViewModel() as MembershipTypeViewModel
                : MembershipMapper.ToVM(ef.MembershipTypes.Find(Encrypter.DecryptToInt(id)), CurrentCulture);
            return View(viewModel);
        }

        [HttpPost]
        public ActionResult EditMembershipType(MembershipTypeViewModel viewModel)
        {
            if(ModelState.IsValid)
            {
                MembershipType originalObject = ef.MembershipTypes.Find(Encrypter.DecryptToInt(viewModel.MembershipTypeID));
                if(originalObject != null)
                {
                    MembershipType newObject = MembershipMapper.ToEF(viewModel);
                    try
                    {
                        ef.Entry(originalObject).CurrentValues.SetValues(newObject);
                        ef.SaveChanges();

                        StoreSuccessMessage(string.Format(SuccessMessages.MembershipTypeUpdated, 
                            CurrentCulture == Culture.en ? originalObject.NameE : originalObject.NameF));
                        return RedirectToAction("ManageMembershipTypes");
                    }
                    catch (Exception)
                    {
                        return RedirectToAction("Index", "Error");
                    }
                }

                StoreModelState("", ModelStateErrors.GeneralError);
                return RedirectToAction("ManageMembershipTypes");
            }

            StoreViewModel(viewModel);
            StoreModelState();
            return RedirectToAction("EditMembershipType", new { id = viewModel.MembershipTypeID }); 
        }

        public ActionResult ReactivateMembershipType(string id)
        {
            int deId = Encrypter.DecryptToInt(id);
            if(deId == (short)MembershipTypes.Administrator 
                && deId != (short)MembershipTypes.NonMember 
                && deId != (short)MembershipTypes.BECOR)
            {
                return RedirectToAction("Index", "Error");
            }

            MembershipType membershipType = ef.MembershipTypes.Find(deId);
            MembershipTypeViewModel viewModel = MembershipMapper.ToVM(membershipType, CurrentCulture);
            if (membershipType != null && membershipType.IsActive == false)
            {
                try
                {
                    viewModel.IsActive = true;
                    ef.Entry(membershipType).CurrentValues.SetValues(MembershipMapper.ToEF(viewModel));
                    ef.SaveChanges();

                    StoreSuccessMessage(string.Format(SuccessMessages.MembershipReactivated, CurrentCulture == Culture.en ? viewModel.NameE : viewModel.NameF));
                    return RedirectToAction("ManageMembershipTypes");
                }
                catch (Exception)
                {
                    return RedirectToAction("Index", "Error");
                }
            }
            StoreModelState("", ModelStateErrors.GeneralError);
            return RedirectToAction("ManageMembershipTypes");   
        }

        public ActionResult DeactivateMembershipType(string id)
        {
            int deId = Encrypter.DecryptToInt(id);
            if (deId == (short)MembershipTypes.Administrator
                && deId != (short)MembershipTypes.NonMember
                && deId != (short)MembershipTypes.BECOR)
            {
                return RedirectToAction("Index", "Error");
            }

            MembershipType membershipType = ef.MembershipTypes.Find(deId);
            MembershipTypeViewModel viewModel = MembershipMapper.ToVM(membershipType, CurrentCulture);
            if (membershipType != null && membershipType.IsActive == true)
            {
                try
                {
                    viewModel.IsActive = false;
                    ef.Entry(membershipType).CurrentValues.SetValues(MembershipMapper.ToEF(viewModel));
                    ef.SaveChanges();

                    StoreSuccessMessage(string.Format(SuccessMessages.MembershipDeactivated, CurrentCulture == Culture.en ? viewModel.NameE : viewModel.NameF));
                    return RedirectToAction("ManageMembershipTypes");
                }
                catch (Exception)
                {
                    return RedirectToAction("Index", "Error");
                }
            }
            StoreModelState("", ModelStateErrors.GeneralError);
            return RedirectToAction("ManageMembershipTypes");   
        }
        #endregion   

        #region Bulk Email

        public ActionResult BulkEmailAddIndividual(string listId, string email)
        {
            // TODO: FUTURE DEVELOPER - MailChimp - add an interface for adding multiple emails at once (comma delimited multilinebox or similar perhaps)
            var mailChimp = new MailChimpManagement(CurrentCulture);
            mailChimp.AddEmailToList(listId, email);

            // TODO: FUTURE DEVELOPER - MailChimp - after this action, should direct back to edit individual with the email searched 
            //ViewBag.email = email;
            return RedirectToAction("BulkEmailEditIndividual");
        }

        public ActionResult BulkEmailCampaigns()
        {
            var viewModel = new BulkEmailCampaignsViewModel();
            var mailChimp = new MailChimpManagement(CurrentCulture);
            viewModel.EmailCampaigns = mailChimp.GetEmailCampaigns();

            return View(viewModel);
        }

        public ActionResult BulkEmailCreateCampaign(string id)
        {
            var viewModel = new BulkEmailCreateCampaignViewModel();
            viewModel.CampaignTypes = CampaignTypes();
            viewModel.Announcements = AnnouncementList();
            viewModel.Seminars = SeminarList();

            return View(viewModel);
        }

        [HttpPost]
        public ActionResult BulkEmailCreateCampaign(BulkEmailCreateCampaignViewModel viewModel)
        {
            var mailChimp = new MailChimpManagement(CurrentCulture);
            SelectList listOfEmailLists = mailChimp.GetListOfEmailLists();
            SelectListItem emailList = listOfEmailLists.First();
            ViewBag.ListEmailLists = listOfEmailLists;

            var editViewModel = new BulkEmailCampaignViewModel
            {
                CreateOptions = new CampaignCreateOptions
                {
                    FromEmail = ConfigurationManager.AppSettings["applicationFromAddress"],
                    FromName = "BECOR",
                    ListId = emailList.Value,
                    Subject = "Default Subject"
                },
                CreateContent = new CampaignCreateContent(),
                SegmentOptions = new CampaignSegmentOptions(),
                TypeOptions = new CampaignTypeOptions()
            };

            var campaign = (EmailCampaign)Enum.Parse(typeof(EmailCampaign), viewModel.CampaignType);

            switch (campaign)
            {
                case EmailCampaign.Announcement:
                    editViewModel = MapNewAnnouncementCampaign(editViewModel, viewModel.Announcement);
                    break;
                case EmailCampaign.Seminar:
                    editViewModel = MapNewSeminarCampaign(editViewModel, viewModel.Seminar);
                    break;
            }

            try
            {

                string campaignId = mailChimp.CreateEmailCampaign(editViewModel);
                return RedirectToAction("BulkEmailEditCampaign", new { id = campaignId });
            }
            catch
            {
                return View("Error");
            }

            return View();
        }

        public ActionResult BulkEmailDeleteCampaign(string id)
        {
            // TODO: FUTURE DEVELOPER - MailChimp - DeleteCampaign - direct to another page to confirm deletion then delete campaign
            var mailChimp = new MailChimpManagement(CurrentCulture);
            mailChimp.DeleteEmailCampaign(id);
            ViewBag.BulkEmailMessage = "Email Campaign Deleted";
            return RedirectToAction("BulkEmailCampaigns");
        }

        public ActionResult BulkEmailEditCampaign(string id)
        {
            var mailChimp = new MailChimpManagement(CurrentCulture);
            ViewBag.ListEmailLists = mailChimp.GetListOfEmailLists();
            var campaign = mailChimp.GetEmailCampaignById(id);

            var viewModel = new BulkEmailCampaignViewModel
            {
                CreateOptions = new CampaignCreateOptions
                {
                    Title = campaign.Title,
                    ListId = campaign.ListId,
                    FromEmail = campaign.FromEmail,
                    FromName = campaign.FromName,
                    Subject = campaign.Subject
                },
                CreateContent = new CampaignCreateContent
                {
                    HTML = mailChimp.GetCampaignHtmlContent(id)
                },
                SegmentOptions = campaign.SegmentOpts,
                TypeOptions = new CampaignTypeOptions(),
                CampaignId = id
            };

            return View(viewModel);
        }

        [ValidateInput(false)]
        [HttpPost]
        public ActionResult BulkEmailEditCampaign(BulkEmailCampaignViewModel viewModel)
        {
            try
            {
                var mailChimp = new MailChimpManagement(CurrentCulture);
                mailChimp.UpdateEmailCampaign(viewModel);

                return RedirectToAction("BulkEmailCampaigns");
            }
            catch
            {
                return View("Error");
            }
        }

        public ActionResult BulkEmailEditIndividual()
        {
            var mailChimp = new MailChimpManagement(CurrentCulture);
            ViewBag.ListEmailLists = mailChimp.GetListOfEmailListsForIndividualAdditions(new[]{" "});
            var viewModel = new IndividualEmailListsViewModel
            {
                EmailLists = new List<ListForEmail>(),
                EmailParameter = new EmailParameter()
            };
            return View(viewModel);
        }

        [HttpPost]
        public ActionResult BulkEmailEditIndividual(IndividualEmailListsViewModel viewModel)
        {
            var mailChimp = new MailChimpManagement(CurrentCulture);
            try
            {
                bool valid = !String.IsNullOrEmpty(viewModel.EmailParameter.Email);
                if (valid)
                {
                    valid = Regex.IsMatch(viewModel.EmailParameter.Email,
                        @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                        @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                        RegexOptions.IgnoreCase);
                    if (valid)
                    {
                        var listsForEmail = mailChimp.GetListsforEmail(viewModel.EmailParameter);
                        var inListIds = new string[listsForEmail.Count];
                        int i = 0;
                        foreach (var list in listsForEmail)
                        {
                            inListIds[i] = list.Id;
                            i++;
                        }

                        viewModel.EmailLists = listsForEmail;
                        ViewBag.ListEmailLists = mailChimp.GetListOfEmailListsForIndividualAdditions(inListIds);
                        return View(viewModel);
                    }
                }

                if (!valid)
                {
                    ViewBag.BulkEmailMessage = "Invalid email - " + viewModel.EmailParameter.Email;
                    viewModel.EmailLists = new List<ListForEmail>();
                    viewModel.EmailParameter.Email = string.Empty;
                }
            }
            catch (MailChimpAPIException)
            {
                ViewBag.ListEmailLists = mailChimp.GetListOfEmailListsForIndividualAdditions(new[] { " " });
                viewModel.EmailLists = new List<ListForEmail>();
            }
            catch (Exception)
            {
                viewModel.EmailLists = new List<ListForEmail>();
            }


            return View(viewModel);
        }

        public ActionResult BulkEmailLists()
        {
            var viewModel = new BulkEmailListsViewModel();
            var mailChimp = new MailChimpManagement(CurrentCulture);
            viewModel.EmailLists = mailChimp.GetEmailLists();

            return View(viewModel);
        }

        public ActionResult BulkEmailOverview()
        {
            return View();
        }

        [HttpPost]
        public ActionResult BulkEmailOverview(string holder)
        {
            var fullList = ef.UserProfiles;

            foreach (var member in fullList)
            {
                var sadgasdg = member;
                UpdateMailChimpMembershipLists(member);
            }

            ViewBag.BulkEmailMessage = "MailChimp email distribution lists have been updated";

            return View();
        }

        public ActionResult BulkEmailRemoveIndividual(string listId, string email)
        {
            var mailChimp = new MailChimpManagement(CurrentCulture);
            mailChimp.RemoveEmailFromList(listId, email);

            // TODO: FUTURE DEVELOPER - MailChimp - after this action, should direct back to edit individual with the email searched 
            //ViewBag.email = email;
            return RedirectToAction("BulkEmailEditIndividual");
        }

        public ActionResult BulkEmailSendCampaign(string id)
        {
            var mailChimp = new MailChimpManagement(CurrentCulture);
            bool successful = mailChimp.SendEmailCampaign(id);

            ViewBag.BulkEmailMessage = (successful) ? "Your email campaign is scheduled to be sent" : "There was an error sending your campaign. Please try again";

            return RedirectToAction("BulkEmailCampaigns");
        }

        public ActionResult BulkEmailViewList(string id)
        {
            var mailChimp = new MailChimpManagement(CurrentCulture);
            var viewModel = AdministrationBulkEmailMapper.EditBulkEmailList(mailChimp.GetEmailListById(id));
            return View(viewModel);
        }

        #endregion

        #region Custom SelectLists
        
        private SelectList CampaignTypes()
        {
            var enumList = new List<SelectListItem>();

            foreach (var suit in Enum.GetValues(typeof(EmailCampaign)))
            {
                var sli = new SelectListItem
                {
                    Value = suit.ToString(),
                    Text = suit.ToString()
                };
                enumList.Add(sli);
            }

            SelectList returnList = new SelectList(enumList, "Value", "Text", "Other");

            return returnList;
        }

        private SelectList SeminarList()
        {
            var dbTable = ef.Events.Select(x => x).ToList().OrderBy(x => x.EndDate);
            var itemlist = new List<SelectListItem>();

            foreach (var record in dbTable)
            {
                var sli = new SelectListItem
                {
                    Text = record.NameE,
                    Value = record.EventID.ToString()
                };
                itemlist.Add(sli);
            }

            SelectList returnList = new SelectList(itemlist, "Value", "Text");

            return returnList;
        }

        private SelectList AnnouncementList()
        {
            var dbTable = ef.Announcements.Select(x => x).ToList().OrderBy(x => x.CreatedDate);
            var itemlist = new List<SelectListItem>();

            foreach (var record in dbTable)
            {
                var sli = new SelectListItem
                {
                    Text = record.SubjectE,
                    Value = record.AnnouncementID.ToString()
                };
                itemlist.Add(sli);
            }

            SelectList returnList = new SelectList(itemlist, "Value", "Text");

            return returnList;
        }

        #endregion

        #region MailChimp Mapping

        private BulkEmailCampaignViewModel MapNewSeminarCampaign(BulkEmailCampaignViewModel viewModel, string seminarId)
        {
            // TODO: FUTURE DEVELOPER - MailChimp - Requires testing
            int semId = Int32.Parse(seminarId);
            var seminar = ef.Events.Where(x => x.EventID == semId).ToList().FirstOrDefault();
            var address = ef.EventAddresses.Where(x => x.EventAddressID == seminar.EventAddressID).ToList().FirstOrDefault();
            var presenters = ef.EventPresenters.Where(x => x.EventID == seminar.EventID).ToList();
            var costs = ef.EventCosts.Where(x => x.EventID == seminar.EventID).ToList();

            string HTML;
            HTML = "<h2>" + seminar.NameE + "</h2>" + Environment.NewLine;
            HTML += "<strong>Start: </strong>" + seminar.StartDate.ToShortDateString() + " - " + seminar.StartTime;
            HTML += "<strong>End: </strong>" + seminar.EndDate.ToShortDateString() + " - " + seminar.EndTime + Environment.NewLine;
            HTML += "<strong>Capacity: </strong>" + seminar.Capacity + Environment.NewLine + Environment.NewLine;

            // event address object
            if (address != null)
                HTML += "<strong>Address: </strong>" + address.Name + ", " + address.Street + ", " + address.CityType.NameE + ", " + address.ProvinceType.NameE
                    + Environment.NewLine + Environment.NewLine;

            // presenter object(s)
            if (presenters != null)
                HTML += "<strong>Presenter:</strong>" + Environment.NewLine + Environment.NewLine;
            foreach (var presenter in presenters)
            {
                HTML += presenter.Name + ", " + presenter.CompanyE + Environment.NewLine;
                HTML += presenter.DetailsE + Environment.NewLine + Environment.NewLine;
            }

            HTML += seminar.ContentE + Environment.NewLine + Environment.NewLine;

            // event cost objects
            if (costs != null)
                HTML += "<strong>Costs:</strong>" + Environment.NewLine + Environment.NewLine;
            foreach (var cost in costs)
            {
                HTML += cost.MembershipType.NameE + Environment.NewLine;
                HTML += cost.EventCost1 + "(+$" + cost.EventHST + ")" + Environment.NewLine + Environment.NewLine;
            }

            viewModel.CreateContent.HTML = HTML;
            viewModel.CreateOptions.Subject = seminar.NameE;

            return viewModel;
        }

        private BulkEmailCampaignViewModel MapNewAnnouncementCampaign(BulkEmailCampaignViewModel viewModel, string announcementId)
        {
            int annId = Int32.Parse(announcementId);
            var announcement = ef.Announcements.Where(x => x.AnnouncementID == annId).ToList().FirstOrDefault();

            viewModel.CreateContent.HTML = announcement.ContentE;
            viewModel.CreateOptions.Subject = announcement.SubjectE;
            
            return viewModel;
        }

        private BulkEmailCampaignViewModel MapNewRenewalCampaign(BulkEmailCampaignViewModel viewModel)
        {

            return viewModel;
        }


        #endregion

        private void UpdateMailChimpMembershipLists(UserProfile userProfile)
        {
            // TODO: FUTURE DEVELOPER - MailChimp - this method currently adds BECOR email records to a MailChimp list, it does not remove them...

            var mailChimp = new MailChimpManagement(CurrentCulture);

            // each list's ID
            string listStudent = "45aab10de1";
            string listNonMembers = "96012d807d";
            string listIndividualMembers = "34a65abfed";
            string listCorporateMembers = "f73af98a93";
            string listPremiumCorporateMembers = "a3e2001302";
            string listPremiumPlusCorporateMember = "380161ea52";
            string listAllCorporateMembers = "97abc2a635";
            string listAllBecorMembers = "711c399e73";
            string listAdministratorMembers = "b3b8dbc724";
            string listBECORMembershipCategory = "38892e8d2a";
            string listUnrenewedMembers = "2b84be44e8";
            
            // gets list IDs for each list member is in
            var listsForEmail = mailChimp.GetListsforEmail(new EmailParameter { Email = userProfile.EmailAddress});
            var inListIds = new string[listsForEmail.Count];
            int i = 0;
            foreach (var list in listsForEmail)
            {
                inListIds[i] = list.Id;
                i++;
            }

            if (userProfile.IsSubscribedToMailingList)
            {
                switch (userProfile.MembershipTypeID)
                {
                    case 1: // Non-Member
                        if (!inListIds.Equals(listNonMembers))
                            mailChimp.AddEmailToList(listNonMembers, userProfile.EmailAddress);
                        break;

                    case 2: // Student
                        if (!inListIds.Equals(listStudent))
                            mailChimp.AddEmailToList(listStudent, userProfile.EmailAddress);
                        goto case 99;

                    case 3: // Individual
                        if (!inListIds.Equals(listIndividualMembers))
                            mailChimp.AddEmailToList(listIndividualMembers, userProfile.EmailAddress);
                        goto case 99;

                    case 4: // Corporate
                        if (!inListIds.Equals(listCorporateMembers))
                            mailChimp.AddEmailToList(listCorporateMembers, userProfile.EmailAddress);
                        goto case 90;

                    case 5: // Premium Corporate
                        if (!inListIds.Equals(listPremiumCorporateMembers))
                            mailChimp.AddEmailToList(listPremiumCorporateMembers, userProfile.EmailAddress);
                        goto case 90;

                    case 6: // Premium Plus Corporate
                        if (!inListIds.Equals(listPremiumPlusCorporateMember))
                            mailChimp.AddEmailToList(listPremiumPlusCorporateMember, userProfile.EmailAddress);
                        goto case 90;

                    case 7: // BECOR
                        if (!inListIds.Equals(listBECORMembershipCategory))
                            mailChimp.AddEmailToList(listBECORMembershipCategory, userProfile.EmailAddress);
                        goto case 99;

                    case 8: // Administrator
                        if (!inListIds.Equals(listAdministratorMembers))
                            mailChimp.AddEmailToList(listAdministratorMembers, userProfile.EmailAddress);
                        goto case 99;
                        
                    case 90: // All Corporate Members
                        if (!inListIds.Equals(listAllCorporateMembers))
                            mailChimp.AddEmailToList(listAllCorporateMembers, userProfile.EmailAddress);
                        goto case 99;

                    case 99: // All members
                        if (!inListIds.Equals(listAllBecorMembers))
                            mailChimp.AddEmailToList(listAllBecorMembers, userProfile.EmailAddress);
                        break;
                }

                bool renewed = (BECORConstants.MEMBERSHIP_EXPIRATION > userProfile.MembershipExpiration && userProfile.MembershipExpiration != DateTime.MinValue);

                // TODO: FUTURE DEVELOPER - MailChimp - needs testing
                // member is not renewed and not a corporate underling
                if (!renewed && userProfile.MembershipTypeID != 4)
                {
                    if (!inListIds.Equals(listUnrenewedMembers))
                        mailChimp.AddEmailToList(listUnrenewedMembers, userProfile.EmailAddress);
                }
            }
            
            // if user has selected to not receive emails
            else
            {
                foreach (var list in inListIds)
                {
                    mailChimp.RemoveEmailFromList(list, userProfile.EmailAddress);
                }
            }
        }

            public ActionResult ExportCSV(object sender, EventArgs e)
        {
            
            string constr = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT EmailAddress FROM UserProfiles"))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter())
                    {
                        cmd.Connection = con;
                        sda.SelectCommand = cmd;
                        using (DataTable dt = new DataTable())
                        {
                            sda.Fill(dt);

                            //Build the CSV file data as a Comma separated string.
                            string csv = string.Empty;

                            foreach (DataColumn column in dt.Columns)
                            {
                                //Add the Header row for CSV file.
                                csv += column.ColumnName + ',';
                            }

                            //Add new line.
                            csv += "\r\n";

                            foreach (DataRow row in dt.Rows)
                            {
                                foreach (DataColumn column in dt.Columns)
                                {
                                    //Add the Data rows.
                                    csv += row[column.ColumnName].ToString().Replace(",", ";") + ',';
                                }

                                //Add new line.
                                csv += "\r\n";
                            }

                            //Download the CSV file.
                            Response.Clear();
                            Response.Buffer = true;
                            Response.AddHeader("content-disposition", "attachment;filename=SqlExport.csv");
                            Response.Charset = "";
                            Response.ContentType = "application/text";
                            Response.Output.Write(csv);
                            Response.Flush();
                            Response.End();
                            return RedirectToAction("BulkEmailOverview");
                        }
                    }
                }
            }
        }


    }
        }
