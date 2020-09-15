using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Security.AccessControl;
using System.Security.Principal;
using AejWNetworkDrive;

namespace createReadme
{
	class Program
	{
		static void Main(string[] args)
		{
			string[] drives = Directory.GetLogicalDrives();
			NetworkDrive networkDrive = new NetworkDrive();
			networkDrive.Persistent = true;
			networkDrive.SaveCredentials = true;
			networkDrive.Force = true;
			networkDrive.LocalDrive = "Z:";
			networkDrive.ShareName = "\\\\10.112.139.3\\c$\\FILERDATEN\\NTFSPhase3";

			try { networkDrive.UnMapDrive(); }
			catch { }
			try
			{
				string username = "NRWBANKDEV\\OIMNTFS";
				string password = "HbLjSEsgv/9ctvj2pYosOJT7UPVpid3qdJP5RPBVbG8=";
				networkDrive.MapDrive(username, password);

				Console.WriteLine("Target {0} mapped to drive {1}.", networkDrive.ShareName, networkDrive.LocalDrive);
			}
			catch (Exception e)
			{
				Console.WriteLine("unable to map drive {0} to {1}", networkDrive.LocalDrive, networkDrive.ShareName);
				Console.WriteLine("{0}", e.ToString());
			}

			SqlConnection conn = new SqlConnection("Data Source=10.112.139.4;Initial Catalog=OIMV8DEV;User Id = oimntfsdbo; Password = HbLjSEsgv/9ctvj2pYosOJT7UPVpid3qdJP5RPBVbG8=");
			string SQL = @"
				SELECT
					UID_CCC_NTFS_Directory AS [UID_Folder],
					CCC_Name as FolderName,
					CCC_Description as Description,
					CCC_FullPath as FullPath,
					CCC_UNCPath as UNCPath,
					CASE CCC_RiskIndex WHEN 0 THEN 'Normal' ELSE 'Admin Spezial' END as ITSDGroup,
					CCC_DateCreation as DateCreated,
					CCC_Type as FolderType,
					CCC_IsInheritanceBreak as isProtected,
					CCC_ShortName as FolderShortName,
					CCC_IsRoot as isRoot,
					CASE CCC_XOrigin WHEN 0 THEN 'Altbestand' ELSE 'OneIM' END as Origin,
					o1.InternalName + ' (' + o1.CentralAccount + ')' as DataOwner1,
					o2.InternalName + ' (' + o2.CentralAccount + ')' as DataOwner2,
					ro.SAMAccountName as ROGroup,
					rw.SAMAccountName as RWGroup
				  FROM [CCC_NTFS_Directory]
				  JOIN Person o1 ON o1.UID_Person = CCC_UID_DataOwner1
				  LEFT OUTER JOIN Person o2 ON o2.UID_Person= CCC_UID_DataOwner2
				  LEFT OUTER JOIN ADSGroup ro ON ro.UID_ADSGroup = CCC_UID_ADSGroupRO
				  LEFT OUTER JOIN ADSGroup rw ON rw.UID_ADSGroup = CCC_UID_ADSGroupRW
				  WHERE CCC_HasEntitlements = 1
				";

			if (args.Length > 0)
			{
				SQL += " AND UID_CCC_NTFS_Directory = '" + args[0] + "'";
			}
			DataTable NodesTable = new DataTable("NodesTable");
			SqlDataAdapter NodesAdapter = new SqlDataAdapter(SQL, conn);

			NodesAdapter.Fill(NodesTable);
			foreach (DataRow node in NodesTable.Rows)
			{
				string targetpath = node["UNCPath"].ToString();
				if (!Directory.Exists(targetpath))
				{
					Console.WriteLine("Directory {0} does not exist.", targetpath);
					continue;
				}

				string fileName = targetpath + "\\readme.html";
				List<string> lines = new List<string>();
				lines.Add("<html><!doctype html><head><meta charset=\"utf-8\">");
				lines.Add("<title>" + node["FullPath"] + "</title></head>");
				lines.Add(" ");
				string css = @"<style> 
					h1 { 
						font-family: Arial, sans-serif;
						font-size: large;
						background-color: #024DA1; color: white;
						padding-top: 6pt;
						padding-bottom: 6pt;
						padding-left: 6pt;
						margin-top: 36pt;
						margin-bottom: 18pt;
					} 
					h2 { 
						font-family: Arial, sans-serif;
						font-size: larger;
						font-color: #024DA1
						padding-top: 3pt;
						padding-bottom: 3pt;
						margin-top: 18pt;
						margin-bottom: 6pt;
					} 
					h3 { 
						font-family: Arial, sans-serif;
						font-color: #024DA1
						font-size: medium;
						padding-top: 3pt;
						padding-bottom: 3pt;
						margin-top: 12pt;
						margin-bottom: 6pt;
					} 
					th, caption { 
						font-family: Arial, sans-serif;
						font-weight: bold;
						text-align: left;
						background-color: #024DA1; color: white;
					}
					td { 
						font-family: Arial, sans-serif;
						text-align: left;
					}
					p, li, blockquote { 
						font-family: Arial, sans-serif;
						text-align: left; 
					}
					blockquote { 
						font-size: smaller;
					}
					table {
						caption-side: top;
						border-color: #024DA1;
					}
					href, a:any-link {
						color: #024DA1;
					}
					th, td, caption {
						border: 1px solid #024DA1;
						padding-top: 3pt;
						padding-bottom: 3pt;
						padding-left: 6pt;
					}
				</style>";
				lines.Add(css);
				lines.Add("<img src=\"https://www.nrwbank.de/export/resources/img/master/logo_de.png\" alt=NRWBank Intranet></head><body>");
				lines.Add("<h1>Verzeichnisinformationen und Berechtigungskonzept</h1>");
				lines.Add("<p>Dieses Verzeichnis <b>" + node["FullPath"] + "</b> wird durch das zentrale IAM-System verwaltet. Anträge zur Änderung der Berechtigung oder Löschung des Verzeichnisses können <a href=https://oimportal.nrwbanki.de/IdentityManager/>im OIM Portal</a> gestellt werden.</p>");
				if (node["Origin"].ToString() == "Altbestand")
				{
					lines.Add("<p>Schreib- und Leseberechtigungen beruhen noch auf der alten Vergabe mit OGITIX. Zur Umstellung stellen Sie bitte einen Antrag auf Änderung der Verzeichnisberechtigungen <a title=\"Link zur direkten Bestellung einer Berechtigungsänderung im Verzeichnis\" href=https://oimportal.nrwbanki.de/IdentityManager/deeplink/Verzeichnisberechtigungen>im OneIM Portal</a>.</p>");
				}
				else
				{
					lines.Add("<p>Schreib- und Leseberechtigungen werden durch Mitgliedschaft in verwalteten Gruppen erteilt. Administrative Berechtigungen hat ausschließlich das zentrale IAM-System über den technischen Benutzer 'OIMNTFS'. Eine Berechtigungsänderung ist so nur durch das System möglich. Sonderberechtigungen sind nicht vergeben. ");
					if (node["isProtected"].ToString() == "true")
					{
						lines.Add("Die Vererbung von Berechtigungen aus übergeordneten Verzeichnissen ist an diesem Verzeichnis unterbrochen. Daher wirken übergeordnete Berechtigungen NICHT auf dieses Verzeichnis.");
					} 
					else
					{
						lines.Add("Die Vererbung von Berechtigungen aus übergeordneten Verzeichnissen werden an dieses Verzeichnis durchgereicht. Daher wirken übergeordnete Berechtigungen auch auf dieses Verzeichnis.");
					}
					lines.Add("Die Berechtigungen werden von diesem Verzeichnis an darin enthaltene Unterverzeichnisse und Dateien vererbt.<p>");

					lines.Add("<h2>Schreibberechtigungen</h2>");
					lines.Add("<p>Schreibberechtigungen auf dieses Verzeichnis werden durch Mitgliedschaft in der Gruppe <a title=\"Link zur direkten Bestellung einer Schreibberechtigung für dieses Verzeichnis\" href=https://oimportal.nrwbanki.de/IdentityManager/deeplink/" + node["RWGroup"] + ">" + node["RWGroup"] + "</a> erworben.</p>");
					lines.Add("<p>Folgende Benutzer und Gruppen haben schreibenden Zugriff:</p>");
					lines.Add("<ul>");
					lines.Add("<li>TBD</li>");
					lines.Add("</ul>");
					lines.Add("<h2>Leseberechtigungen</h2>");
					lines.Add("<p>Leseberechtigungen auf dieses Verzeichnis werden durch Mitgliedschaft in der Gruppe <a title=\"Link zur direkten Bestellung einer Leseberechtigung für dieses Verzeichnis\" href=https://oimportal.nrwbanki.de/IdentityManager/deeplink/" + node["ROGroup"] + ">" + node["ROGroup"] + "</a> erworben.</p>");
					lines.Add("<p>Folgende Benutzer und Gruppen haben schreibenden Zugriff:</p>");
					lines.Add("<ul>");
					lines.Add("<li>TBD</li>");
					lines.Add("</ul>");
				}
				lines.Add("<p>Dateneigner und damit Genehmiger für Berechtigungsanträge sind " + node["DataOwner1"] + " und " + node["DataOwner2"] + "</p>");

				lines.Add("<h2>Berechtigungsvergabe</h2>");
				lines.Add("<p>Schreibberechtigte Personen können innerhalb dieses Verzeichnisses beliebige Ordnerstrukturen aufbauen. Alle Ordner erben die Berechtigungen aus diesem Verzeichnis. Falls in einem Unterordner spezielle Berechtigungen gesetzt werden müssen, kann dies ebenfalls im OIM-Portal beantragt werden. Der ITSD berät gern über die Nutzung des Portals und die Strukturierung der Ordner.</p>");
				lines.Add("<p>Bitte beachten Sie, dass der ITSD keine Berechtigungen in diesem Verzeichnis mehr vergeben oder ändern kann. Alle Berechtigungen müssen über das OIM-Portal verwaltet werden. Dieses stellt die Einhaltung aller bankaufsichtlichen Anforderungen sicher und ist daher verbindlich für alle Benutzerberechtigungsverwaltungsvorgänge einzusetzen. Vgl. auch BAIT Kapitel 5, Ziffern 23ff. sowie MaRisk AT 4.3.1 Tz. 2, AT 7.2 Tz. 2 und BTO Tz. 9.</p>");
				lines.Add("<br/>");
				lines.Add("<p style = \"font-style: italic; text-align: right; font-size: small; margin: 6pt; white-space: pre-line;\">");
				lines.Add("<b>Quellen:</b>");
				lines.Add("BAIT: <a href=https://www.bafin.de/SharedDocs/Downloads/DE/Rundschreiben/dl_rs_1710_ba_BAIT.pdf>https://www.bafin.de/SharedDocs/Downloads/DE/Rundschreiben/dl_rs_1710_ba_BAIT.pdf</a>");
				lines.Add("MaRisk: <a href=https://www.bafin.de/SharedDocs/Veroeffentlichungen/DE/Rundschreiben/2017/rs_1709_marisk_ba.html>https://www.bafin.de/SharedDocs/Veroeffentlichungen/DE/Rundschreiben/2017/rs_1709_marisk_ba.html</a>");
				lines.Add("</p>");

				lines.Add("<h1>Hinweise zur Verwaltung komplexer Verzeichnisberechtigungen</h1>");
				lines.Add("<p>Bankaufsichtlich sind einige zentrale Forderungen an Verzeichnisberechtigungen zu stellen:</p>");
				lines.Add("<ul>");
				lines.Add("<li>Alle von einem IT-System bereitgestellten Berechtigungen müssen vollständig und nachvollziehbar ableitbar in einem Berechtigungskonzept beschrieben werden (BAIT, Ziffer 24).</li>");
				lines.Add("<li>Alle Verfahren zur Berechtigungsverwaltung müssen durch Genehmigungs- und Kontrollprozesse sicherstellen, dass die Vorgaben des Berechtigungskonzepts eingehalten werden. Dies umfasst auch die Umsetzung des Berechtigungsantrags im Zielsystem (BAIT, Ziffer 26).</li>");
				lines.Add("<li>Alle IT-gestützten Verfahren müssen die Funktionstrennung sicherstellen (MaRisk, BTO Tz. 9).</li>");
				lines.Add("</ul>");
				lines.Add("<p>Wenn Standardstrukturen nicht flexibel genug oder nicht ausreichend sind, erarbeiten Anwender der Bank mit dem IT Service Desk sogenannte komplexe Verzeichnisberechtigungen, die beschreiben, welche Verzeichnisse angelegt und welche Benutzer auf welchen Verzeichnissen in welcher Form berechtigt werden sollen. Mit Einführung des zentralen Verfahrens One Identity Manager werden die komplexen Berechtigungen in diesem Verfahren gebildet, genehmigt und umgesetzt, um so die bankaufsichtlichen Anforderungen nachhaltig zu erfüllen.</p>");

				lines.Add("<h2>Genehmigungsverfahren</h2>");
				lines.Add("<p>Als zentrales Genehmigungsverfahren in der Bank legt das Handbuch 205 in Kapitel 12.11.2.1 unter Abschnitt 2c fest:</p>");
				lines.Add("<blockquote>Ein Berechtigungsantrag, welcher den Berechtigungsumfang erhöht, muss mindestens durch die Führungskraft der Person, für die eine Berechtigung beantragt wird, oder durch den verantwortlichen und im Antragsprozess dokumentierten Dateneigner genehmigt werden.Weitere Genehmigungsinstanzen sind in Abhängigkeit von der Schutzbedarfsfeststellung(-> Kap. 12.4.1) festzulegen. Berechtigungsanträge für Vorstände und Bereichsleiter können alternativ anstelle der Führungskraft auch durch die Informationssicherheit genehmigt werden. Ein Berechtigungsantrag, welcher den Berechtigungsumfang reduziert, kann von dem Benutzer selber, dessen (vorheriger) Führungskraft oder dem verantwortlichen und im Antragsprozess dokumentierten Dateneigner allein genehmigt werden, ohne dass dies im Antragsprozess gesondert festgelegt sein muss. Ein elektronischer Berechtigungsantragsprozess mittels eines Workflow - Tools mit sicherer Protokollierung ist gleichwertig zu einem papierhaften Antragsverfahren(Antragsformular).</blockquote>");

				lines.Add("<h2>Kontrollverfahren</h2>");
				lines.Add("<p>Nach BAIT Ziffer 26 müssen geeignete Kontrollverfahren sicherstellen, dass Einrichtung, Änderung, Deaktivierung oder Löschung von Berechtigungen den Vorgaben des Berechtigungskonzepts folgen. Dies umfasst sowohl die Genehmigung als auch die Umsetzung des Berechtigungsantrags im jeweiligen Zielsystem.</p>");

				lines.Add("<h3>Überprüfung (Rezertifizierung)</h3>");
				lines.Add("<p>Die Genehmigung einer Berechtigung ist an die Aufgabe des Berechtigungsempfängers gebunden. Im Rahmen des Kontrollverfahrens „Genehmigung“ ist regelmäßig zu überprüfen, ob:</p>");
				lines.Add("<ul>");
				lines.Add("<li>der Berechtigungsantrag eine konkrete Aufgabe benennt, zu deren Erfüllung der Berechtigungsempfänger die Berechtigung benötigt,</li>");
				lines.Add("<li>der Berechtigungsempfänger die Aufgabe, zu deren Erfüllung die Berechtigung benötigt wird, tatsächlich immer noch ausübt und ob</li>");
				lines.Add("<li>der Berechtigungsempfänger keine Berechtigungen, die mit der beantragten Berechtigung im Sinne der Funktionstrennung unvereinbar ist.</li>");
				lines.Add("</ul>");
				lines.Add("<p>Diese Kontrolle der Genehmigung („Rezertifizierung“) erfolgt wie im Fachkonzept Rezertifizierung ausgeführt.</p>");

				lines.Add("<h3>SOLL-IST-Abgleich</h2>");
				lines.Add("<p>Ein genehmigter Berechtigungsantrag wird im Zielsystem technisch umgesetzt. Dies erfolgt für NTFS-Berechtigungen durch Eintrag einer Berechtigungsgruppe an einem Verzeichnis.</p>");
				lines.Add("<p>Im Rahmen des Kontrollverfahrens „Umsetzung“ ist regelmäßig zu prüfen, ob:</p>");
				lines.Add("<ul>");
				lines.Add("<li>alle genehmigten Berechtigungen korrekt im jeweiligen Zielsystem eingetragen sind</li>");
				lines.Add("<li>und ob die eingetragenen Berechtigungen genehmigt sind.</li>");
				lines.Add("</ul>");
				lines.Add("<p>Die Kontrolle muss daher</p>");
				lines.Add("<ol type=\"a\">");
				lines.Add("<li>den Bestand der genehmigten Berechtigungen (SOLL) auf Korrektheit und Vollständigkeit der Umsetzung durch Eintragung im NTFS-Dateisystem und</li>");
				lines.Add("<li>den Bestand im NTFS-Dateisystem eingetragener Berechtigungen (IST) auf Vorliegen einer entsprechenden Genehmigung überprüfen.</li>");
				lines.Add("</ol>");
				lines.Add("<p>Diese Kontrolle erfolgt im Rahmen des SOLL-IST-Abgleichs durch Vergleich der im IAM-System hinterlegten, genehmigten Berechtigungen mit den im NTFS-Dateisystem vorhan-denen Berechtigungen. Dabei werden Gruppenmitgliedschaften durch Standardmecha-nismen des OneIM regelmäßig (derzeit ca. einmal stündlich) abgeglichen, die NTFS-Berechtigungen sind über den SOLL-IST-Vergleich abzugleichen. Der SOLL-IST-Abgleich meldet Abweichungen an die Informationssicherheit und den IT Service Desk.</p>");
				lines.Add("<p style = \"font-style: italic; text-align: right; font-size: small; margin: 6pt; white-space: pre-line;\">Quelle: Fachkonzept Einführung Identity & Access Management System OneIM</p>");

				lines.Add("<h1>Technische Informationen (OneIM Datenbank)</h1>");
				lines.Add("<table><caption>Fields from [CCC_NTFS_Directory]</caption>");
				lines.Add("<tr>");

				for (int i=0; i<node.ItemArray.Length; i++)
				{
					lines.Add("<th>" + NodesTable.Columns[i] + "</th>");
					lines.Add("<td>" + node[i] + "</td>");
					lines.Add("</tr><tr>");
				}
				lines.Add("</tr>");
				lines.Add("</body>");
				lines.Add("</html>");
				Console.WriteLine(targetpath);

				if (File.Exists(fileName))
					File.Delete(fileName);
				
				File.WriteAllLines(fileName, lines.ToArray(),Encoding.UTF8);

				SecurityIdentifier oimNTFSSID = new SecurityIdentifier("S-1-5-21-2847306829-1541239473-3213474396-16514");
				FileSecurity fileSecurity = new FileSecurity(fileName, AccessControlSections.Access);
				fileSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null), FileSystemRights.Read, AccessControlType.Allow));
				fileSecurity.AddAccessRule(new FileSystemAccessRule(oimNTFSSID, FileSystemRights.FullControl, AccessControlType.Allow));
				fileSecurity.SetAccessRuleProtection(true, false);
				File.SetAccessControl(fileName, fileSecurity);
				/*
				lines.Clear();
				fileName = targetpath + "\\desktop.ini";

				if (File.Exists(fileName))
				{
					File.SetAttributes(fileName, FileAttributes.Normal);
					File.Delete(fileName);
				}

				lines.Add("[.ShellClassInfo]");
				lines.Add("IconFile=nbfolder.ico");
				lines.Add("IconIndex = 0");
				lines.Add("InfoTip=Folder manager by Identity & Access Management System.");
				File.WriteAllLines(fileName, lines.ToArray(), Encoding.UTF8);
				File.SetAttributes(fileName, FileAttributes.Hidden | FileAttributes.ReadOnly | FileAttributes.System);

				fileSecurity = new FileSecurity(fileName, AccessControlSections.Access);
				fileSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null), FileSystemRights.Read, AccessControlType.Allow));
				fileSecurity.AddAccessRule(new FileSystemAccessRule(oimNTFSSID, FileSystemRights.FullControl, AccessControlType.Allow));
				fileSecurity.SetAccessRuleProtection(true, false);
				File.SetAccessControl(fileName, fileSecurity);

				fileName = targetpath + "\\nbfolder.ico";
				if (!File.Exists(fileName))
					File.Copy("nbfolder.ico", fileName);

				fileSecurity = new FileSecurity(fileName, AccessControlSections.Access);
				fileSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null), FileSystemRights.Read, AccessControlType.Allow));
				fileSecurity.AddAccessRule(new FileSystemAccessRule(oimNTFSSID, FileSystemRights.FullControl, AccessControlType.Allow));
				fileSecurity.SetAccessRuleProtection(true, false);
				File.SetAccessControl(fileName, fileSecurity);

				DirectoryInfo tInfo = new DirectoryInfo(targetpath);
				tInfo.Attributes |= FileAttributes.System;
				*/
			}
		}
	}
}
