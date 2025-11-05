using CoreExtension.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;

namespace ERP.Common
{
	public class MailHelper
	{
		//protected static readonly log4net.ILog mLog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
		protected static readonly ILogger _Log = LogManager.GetCurrentClassLogger();

		public MailHelper() 
		{ 
			this.SmtpServer = "mail.cht-pt.com.tw";
			this.SmtpPort = 25;
			this.MailAccount = String.Empty;
			this.MailPassword = String.Empty;
		}

		public MailHelper(string smtpServer, int smtpPort, string mailAccount = null, string mailPassword = null)
		{
			if(String.IsNullOrWhiteSpace(smtpServer)) { throw new ArgumentNullException("smtpServer"); }
			if(smtpPort <= 0) { throw new ArgumentNullException("smtpPort"); }

			this.SmtpServer = smtpServer;
			this.SmtpPort = smtpPort;
			this.MailAccount = !String.IsNullOrWhiteSpace(mailAccount)
							 ? mailAccount
							 : String.Empty;
			this.MailPassword = !String.IsNullOrWhiteSpace(mailPassword)
							  ? mailPassword
							  : String.Empty;
		}

		public string SmtpServer { get; private set; } 
		public int SmtpPort { get; private set; } 
		public string MailAccount { get; private set; } 
		public string MailPassword { get; private set; } 

		public MailResult SendMail(MailSetting mailSetting)
		{
			try
			{
				MailMessage message = new MailMessage();
				message.From = new MailAddress(mailSetting.FromAdress);
				AddMaillAdresses(message.To, mailSetting.ToAdress);
				AddMaillAdresses(message.CC, mailSetting.CcAdress);
				AddMaillAdresses(message.Bcc, mailSetting.BccAdress);
				message.Priority = mailSetting.Priority;
				message.BodyEncoding = mailSetting.MailEncoding;
				message.Subject = mailSetting.Subject;
				message.Body = mailSetting.MailBody;
				message.IsBodyHtml = mailSetting.IsBodyHtml;
				if(!mailSetting.Attachments.IsNullOrEmpty())
				{
					foreach(var attachment in mailSetting.Attachments)
					{
						message.Attachments.Add(attachment);
					}
				}

				if(mailSetting.Action != null)
				{
					mailSetting.Action(message);
				}

				NetworkCredential networkCredential = new NetworkCredential(this.MailAccount, this.MailPassword);
				SmtpClient smtpClient = new SmtpClient(this.SmtpServer, this.SmtpPort)
				{
					Credentials = networkCredential,
					UseDefaultCredentials = false
				};

				using(smtpClient)
				{
					using(message)
					{
						smtpClient.Send(message);
					}					
				}

				return new MailResult()
				{
					IsSuccess = true
				};
			}
			catch(Exception ex)
			{
				string msg = String.Format("{0}.{1}-Exception:{2}", MethodBase.GetCurrentMethod().ReflectedType, MethodBase.GetCurrentMethod().Name, ex.ToString());
				_Log.Error(ex, msg);
				return new MailResult()
				{
					IsSuccess = false,
					ErrorMessage = ex.ToString()
				};
			}
		}

		private void AddMaillAdresses(MailAddressCollection mailAddresses, string adresses)
		{
			if(String.IsNullOrWhiteSpace(adresses)) { return; }
			var mailAdresses = (from adress in adresses.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
								where !String.IsNullOrWhiteSpace(adress)
								select new MailAddress(adress)
							   ).ToList();
			foreach(var mailAdress in mailAdresses)
			{
				mailAddresses.Add(mailAdress);
			}
		}
	}

	public class MailSetting
	{
		public MailSetting()
		{
			this.Subject = String.Empty;
			this.MailBody = String.Empty;
			this.FromAdress = String.Empty;
			this.ToAdress = String.Empty;
			this.CcAdress = String.Empty;
			this.BccAdress = String.Empty;
			this.MailEncoding = Encoding.UTF8;
			this.Priority = MailPriority.Normal;
			this.IsBodyHtml = true;
			this.Attachments = new List<Attachment>();
		}

		public string Subject { get; set; } 
		public string MailBody { get; set; }
		public string FromAdress { get; set; }
		public string ToAdress { get; set; }
		public string CcAdress { get; set; }
		public string BccAdress { get; set; }
		public Encoding MailEncoding { get; set; }
		public MailPriority Priority { get; set; }
		public bool IsBodyHtml { get; set; }
		public Action<MailMessage> Action { get; set; }
		public List<Attachment> Attachments { get; private set; } 
	}

	public class MailResult
	{
		public MailResult()
		{
			this.ErrorMessage = String.Empty;
		}

		public bool IsSuccess { get; set; }
		public string ErrorMessage { get; set; }
	}

}