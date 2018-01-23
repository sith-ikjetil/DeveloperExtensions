using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using System.Runtime;
using Microsoft.Win32;
using System.Linq;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
namespace DeveloperExtensions.Setup.CustomAction
{
	public class CustomActions
	{
		[CustomAction]
		public static ActionResult RegisterCOMServers(Session session)
		{
			try
			{
				session.Log("Begin RegisterCOMServers");

				//
				// Get Folders
				//				
				string targetFolder32 = Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "IT Software"), "Developer Extensions");
				string targetFolder64 = targetFolder32.Replace("Program Files (x86)", "Program Files");

				//
				// DLL COM SERVERS
				//
				session.Log("Begin Registering DeveloperExtensions.ContextMenu...");
				ProcessStartInfo startInfo0 = new ProcessStartInfo();
				startInfo0.Arguments = " /s DeveloperExtensions.ContextMenu.dll";
				startInfo0.WorkingDirectory = targetFolder64;
				startInfo0.Verb = "runas";
				startInfo0.FileName = "regsvr32";
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo0))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				session.Log("Begin Registering DeveloperExtensions.PropertySheet...");
				ProcessStartInfo startInfo1 = new ProcessStartInfo();
				startInfo1.Arguments = " /s DeveloperExtensions.PropertySheet.dll";
				startInfo1.WorkingDirectory = targetFolder64;
				startInfo1.Verb = "runas";
				startInfo1.FileName = "regsvr32";
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo1))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				session.Log("Begin Registering DeveloperExtensions.ToolTip...");
				ProcessStartInfo startInfo2 = new ProcessStartInfo();
				startInfo2.Arguments = " /s DeveloperExtensions.ToolTip.dll";
				startInfo2.WorkingDirectory = targetFolder64;
				startInfo2.Verb = "runas";
				startInfo2.FileName = "regsvr32";
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo2))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				//
				// EXE COM SERVERS
				//
				session.Log("Begin Registering DeveloperExtensions.ToolTip64Server...");
				ProcessStartInfo startInfo3 = new ProcessStartInfo();
				startInfo3.Arguments = " /RegServer";
				startInfo3.WorkingDirectory = targetFolder64;
				startInfo3.Verb = "runas";
				startInfo3.FileName = Path.Combine(targetFolder64, "DeveloperExtensions.ToolTip64Server.exe");
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo3))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				session.Log("Begin Registering DeveloperExtensions.ToolTip32Server...");
				ProcessStartInfo startInfo4 = new ProcessStartInfo();
				startInfo4.Arguments = " /RegServer";
				startInfo4.WorkingDirectory = targetFolder32;
				startInfo4.Verb = "runas";
				startInfo4.FileName = Path.Combine(targetFolder32, "DeveloperExtensions.ToolTip32Server.exe");
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo4))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				//
				// Start DeveloperExtensions.TaskBarStatusArea.exe
				//
				ProcessStartInfo startInfo5 = new ProcessStartInfo();
				startInfo5.Arguments = string.Empty;
				startInfo5.WorkingDirectory = targetFolder64;
				startInfo5.Verb = "open";
				startInfo5.FileName = Path.Combine(targetFolder64, "DeveloperExtensions.TaskBarStatusArea.exe");
				System.Diagnostics.Process.Start(startInfo5);
			}
			catch (Exception x)
			{
				session.Log(x.ToString());

				return ActionResult.Failure;
			}
			finally
			{
				session.Log("End RegisterCOMServers");
			}

			return ActionResult.Success;
		}

		[CustomAction]
		public static ActionResult UnregisterCOMServers(Session session)
		{
			try
			{
				session.Log("Begin UnregisterCOMServers");

				//
				// Get Folders
				//
				string targetFolder32 = Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "IT Software"), "Developer Extensions");
				string targetFolder64 = targetFolder32.Replace("Program Files (x86)", "Program Files");

				//
				// DLL COM SERVERS
				//
				session.Log("Begin Unregistering DeveloperExtensions.ContextMenu...");
				ProcessStartInfo startInfo0 = new ProcessStartInfo();
				startInfo0.Arguments = " /s /u DeveloperExtensions.ContextMenu.dll";
				startInfo0.WorkingDirectory = targetFolder64;
				startInfo0.Verb = "runas";
				startInfo0.FileName = "regsvr32";
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo0))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				session.Log("Begin Unregistering DeveloperExtensions.PropertySheet...");
				ProcessStartInfo startInfo1 = new ProcessStartInfo();
				startInfo1.Arguments = " /s /u DeveloperExtensions.PropertySheet.dll";
				startInfo1.WorkingDirectory = targetFolder64;
				startInfo1.Verb = "runas";
				startInfo1.FileName = "regsvr32";
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo1))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				session.Log("Begin Unregistering DeveloperExtensions.ToolTip...");
				ProcessStartInfo startInfo2 = new ProcessStartInfo();
				startInfo2.Arguments = " /s /u DeveloperExtensions.ToolTip.dll";
				startInfo2.WorkingDirectory = targetFolder64;
				startInfo2.Verb = "runas";
				startInfo2.FileName = "regsvr32";
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo2))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				//
				// EXE COM SERVERS
				//
				session.Log("Begin Unregistering DeveloperExtensions.ToolTip64Server...");
				ProcessStartInfo startInfo3 = new ProcessStartInfo();
				startInfo3.Arguments = " /UnregServer";
				startInfo3.WorkingDirectory = targetFolder64;
				startInfo3.Verb = "runas";
				startInfo3.FileName = Path.Combine(targetFolder64, "DeveloperExtensions.ToolTip64Server.exe");
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo3))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				session.Log("Begin Unregistering DeveloperExtensions.ToolTip32Server...");
				ProcessStartInfo startInfo4 = new ProcessStartInfo();
				startInfo4.Arguments = " /UnregServer";
				startInfo4.WorkingDirectory = targetFolder32;
				startInfo4.Verb = "runas";
				startInfo4.FileName = Path.Combine(targetFolder32, "DeveloperExtensions.ToolTip32Server.exe");
				using (Process exeProcess = System.Diagnostics.Process.Start(startInfo4))
				{
					exeProcess.WaitForExit();
				}
				session.Log("OK");

				//
				// Stop DeveloperExtensions.TaskBarStatusArea.exe if running
				//
				session.Log("Begin Stopping DeveloperExtensions.TaskBarStatusArea Process");
				Process.GetProcessesByName("DeveloperExtensions.TaskBarStatusArea").ToList().ForEach(p => p.Kill());
				session.Log("OK");
			}
			catch (Exception x)
			{
				session.Log(x.ToString());

				return ActionResult.Failure;
			}
			finally
			{
				session.Log("End UnregisterCOMServers");
			}

			return ActionResult.Success;
		}

		[CustomAction]
		public static ActionResult RegisterFileAssociations(Session session)
		{
			try
			{
				//
				// Remove File Association Handler
				//
				var tree = Registry.ClassesRoot.OpenSubKey("DeveloperExtensions.EncDec_ext");
				if (tree != null)
				{
					tree.Close();
					Registry.ClassesRoot.DeleteSubKeyTree("DeveloperExtensions.EncDec_ext");
				}

				//
				// Remove File Association for .devenc
				//
				tree = Registry.ClassesRoot.OpenSubKey(".devenc");
				if (tree != null)
				{
					tree.Close();
					Registry.ClassesRoot.DeleteSubKeyTree(".devenc");
				}

				//
				// Remove File Association for .devdec
				//
				tree = Registry.ClassesRoot.OpenSubKey(".devdec");
				if (tree != null)
				{
					tree.Close();
					Registry.ClassesRoot.DeleteSubKeyTree(".devdec");
				}

				//
				// Create File Association
				//			
				var skHandler = Registry.ClassesRoot.CreateSubKey("DeveloperExtensions.EncDec_ext");
				skHandler.SetValue(string.Empty, "Developer Extensions Encryption Decryption File");

				var skShellHandler = skHandler.CreateSubKey("shell");
				var skShellOpenHandler = skShellHandler.CreateSubKey("open");
				var skShellCommandHandler = skShellOpenHandler.CreateSubKey("command");
				skShellCommandHandler.SetValue(string.Empty, $@"""{Environment.ExpandEnvironmentVariables("%ProgramFiles%").Replace(" (x86)", string.Empty)}\It Software\Developer Extensions\DeveloperExtensions.EncDec.exe"" ""%1""");

				var skDefaultIcon = skHandler.CreateSubKey("DefaultIcon");
				skDefaultIcon.SetValue(string.Empty, $@"""{Environment.ExpandEnvironmentVariables("%ProgramFiles%").Replace(" (x86)", string.Empty)}\It Software\Developer Extensions\DeveloperExtensions.EncDec.exe"",0");

				//
				// Create File Association for .devenc
				//			
				var skExtEnc = Registry.ClassesRoot.CreateSubKey(".devenc");
				skExtEnc.SetValue(string.Empty, "DeveloperExtensions.EncDec_ext");
				//skExtEnc.SetValue("Content Type", "application/x-devextencdec");

				//var skShellEnc = skExtEnc.CreateSubKey("shell");
				//var skShellOpenEnc = skShellEnc.CreateSubKey("open");
				//var skShellCommandEnc = skShellOpenEnc.CreateSubKey("command");
				//skShellCommandEnc.SetValue(string.Empty, $@"""{Environment.ExpandEnvironmentVariables("%ProgramFiles%").Replace(" (x86)", string.Empty)}\It Software\Developer Extensions\DeveloperExtensions.EncDec.exe"" ""%1""");

				//
				// Create File Association for .devdec
				//			
				var skExtDec = Registry.ClassesRoot.CreateSubKey(".devdec");
				skExtDec.SetValue(string.Empty, "DeveloperExtensions.EncDec_ext");
				//skExtDec.SetValue("Content Type", "application/x-devextencdec");

				//var skShellDec = skExtDec.CreateSubKey("shell");
				//var skShellOpenDec = skShellDec.CreateSubKey("open");
				//var skShellCommandDec = skShellOpenDec.CreateSubKey("command");
				//skShellCommandDec.SetValue(string.Empty, $@"""{Environment.ExpandEnvironmentVariables("%ProgramFiles%").Replace(" (x86)", string.Empty)}\It Software\Developer Extensions\DeveloperExtensions.EncDec.exe"" ""%1""");
			}
			catch (Exception x)
			{
				session.Log(x.ToString());
				return ActionResult.Failure;
			}
			finally
			{
				session.Log("End RegisterFileAssociations");
			}

			return ActionResult.Success;
		}

		[CustomAction]
		public static ActionResult UnregisterFileAssociations(Session session)
		{
			try
			{
				//
				// Remove File Association Handler
				//
				var tree = Registry.ClassesRoot.OpenSubKey("DeveloperExtensions.EncDec_ext");
				if (tree != null)
				{
					tree.Close();
					Registry.ClassesRoot.DeleteSubKeyTree("DeveloperExtensions.EncDec_ext");
				}

				//
				// Remove File Association for .devenc
				//
				tree = Registry.ClassesRoot.OpenSubKey(".devenc");
				if (tree != null)
				{
					tree.Close();
					Registry.ClassesRoot.DeleteSubKeyTree(".devenc");
				}

				//
				// Remove File Association for .devdec
				//
				tree = Registry.ClassesRoot.OpenSubKey(".devdec");
				if (tree != null)
				{
					tree.Close();
					Registry.ClassesRoot.DeleteSubKeyTree(".devdec");
				}
			}
			catch (Exception x)
			{
				session.Log(x.ToString());
				return ActionResult.Failure;
			}
			finally
			{
				session.Log("End UnregisterFileAssociations");
			}

			return ActionResult.Success;
		}
	}
}
