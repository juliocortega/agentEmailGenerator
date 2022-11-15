
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;

Console.WriteLine("Enter the path");
var msgPath = Console.ReadLine();
//D:\EmailGenerator\FW_ Debt-Terms_xlsx.msg
var msg = new MsgReader.Outlook.Storage.Message(msgPath);
var xlsxPath = Path.GetTempFileName();
File.WriteAllBytes(xlsxPath, ((MsgReader.Outlook.Storage.Attachment)msg.Attachments[0]).Data);
var adapter = new OleDbDataAdapter("select [Agent#] from [Agents to Term$]", $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={xlsxPath};Extended Properties=\"Excel 8.0;HDR=YES\"");
var agentNumbers = new DataTable();
adapter.Fill(agentNumbers);

foreach(DataRow row in agentNumbers.Rows)
{
    var agentNumber = row[0].ToString();
    var client = new System.Net.Mail.SmtpClient(ConfigurationManager.AppSettings["MailHost"]);
    var message = new System.Net.Mail.MailMessage("jcesarortega85@gmail.com", "mktsales@aatx.com", $"agt={agentNumber}, ADTR 22 Level 3",);
    client.Send
    Debugger.Break();
}
Console.WriteLine("Finish");
