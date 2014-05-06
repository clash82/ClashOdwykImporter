// Clash Odwyk Importer
// importer tłumaczeń biblijnych dla programu Odwyk (http://program.odwyk.com)
// (importer tłumaczeń biblijnych ze strony http://biblia.info.pl)
// autor: Rafał Toborek
// http://toborek.info

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.ComponentModel;
using Odwyk.Common;

namespace clash_importer
{
    public class clash_silnik: IBibleImporter
    {
        private ImporterStatus iStatus = ImporterStatus.NotInitialized;
        private ImporterAvailableTranslation[] iTlumaczenia;

        public string ImporterName
        {
            get { return "Clash Odwyk Importer (import z http://biblia.info.pl)"; }
        }

        public string ImporterDescription
        {
            get { return "wersja: " + this.WersjaWtyczki + " (strona domowa: http://toborek.info)"; }
        }

        public ImporterAvailableTranslation[] AvailableTranslations
        {
            get { return this.iTlumaczenia; }
        } 

        public ImporterStatus Status
        {
            get { return this.iStatus; }
        }

        public string ImportTranslation
        {
            set;
            get;
        }

        // aktualnie importowana księga (zmieniane w trakcie importu na bieżąco)
        public string ImportBook
        {
            set;
            get;
        }

        // aktualizacja postępu podczas importowania (zmienianie na bieżąco, od 0 do 100)
        public int ImportProgressPercent
        {
            get;
            set;
        }

        // jeśli wystąpił błąd, to ustwiamy opis tutaj
        public string ImportError
        {
            set;
            get;
        }

        // zdarzenie błędu podczas importu
        public Exception ImportException
        {
            set;
            get;
        }

        // pliku logu
        public string ImportLogFile
        {
            set;
            get;
        }

        // biblia gotowa do użytku (musi być ustawiona po zakończeniu importu, w przeciwnym wypadku ma być pusta - po odpaleniu zdarzenia ImportFinished)
        public BibleTranslation ImportedBible
        {
            set;
            get;
        }

        // domyślny User-Agent wysyłany podczas zapytań do serwera
        private string UserAgent
        {
            get { return "Clash Odwyk Importer (+http://toborek.info)"; }
        }

        // wersja wtyczki
        private string WersjaWtyczki
        {
            get { return "1.2.2.0"; }
        }

        // lista przechowująca informacje o tłumaczeniach (na własny, późniejszy użytek)
        private SortedList<string, string> lsTlumaczenia = new SortedList<string, string>();

        // lista ksiąg do pobrania
        private List<string> lsKsiegi
        {
            get
            {
                List<string> ls = new List<string>
                {
                    // "Tob|1", wtórnokanoniczna
                    // "Jdt|1", wtórnokanoniczna
                    // "1Mch|1", wtórnokanoniczna
                    // "2Mch|1", wtórnokanoniczna
                    // "Mdr|1", wtórnokanoniczna
                    // "Bar|1", wtórnokanoniczna
                    // "Syr|1", wtórnokanoniczna

                    // Stary Testament
                    "Gen|50",
                    "Ex|40",
                    "Lev|27",
                    "Nu|36",
                    "Deu|34",
                    "Joz|24",
                    "Sdz|21",
                    "Rut|4",
                    "1Sam|31",
                    "2Sam|24",
                    "1Krl|22",
                    "2Krl|25",
                    "1Krn|29",
                    "2Krn|36",
                    "Ezd|10",
                    "Neh|13",                    
                    "Est|10",
                    "Hi|42",
                    "Ps|150",
                    "Prz|31",
                    "Koh|12",
                    "Pnp|8",
                    "Iz|66",
                    "Jer|52",
                    "Lam|5",
                    "Ez|48",
                    "Dan|12",
                    "Oz|14",
                    "Joel|3",
                    "Am|9",
                    "Ab|1",
                    "Jon|4",
                    "Mi|7",
                    "Na|3",
                    "Hab|3",
                    "Sof|3",
                    "Ag|2",
                    "Zach|14",
                    "Mal|4",

                    // Nowy Testament
                    "Mt|28",
                    "Mk|16",
                    "Luk|24",
                    "J|21",
                    "Dz|28",
                    "Rz|16",
                    "1Kor|16",
                    "2Kor|13",
                    "Ga|6",
                    "Ef|6",
                    "Fil|4",
                    "Kol|4",
                    "1Tes|5",
                    "2Tes|3",
                    "1Tym|6",
                    "2Tym|4",
                    "Tyt|3",
                    "Flm|1",
                    "Heb|13",
                    "Jak|5",
                    "1P|5",
                    "2P|3",
                    "1J|5",
                    "2J|1",
                    "3J|1",
                    "Jud|1",
                    "Ap|22"
                };
                return ls;
            }
        }

        private string HttpPost(string uri, string parameters)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(uri);
            webRequest.ContentType = "application/x-www-form-urlencoded";
            webRequest.UserAgent = this.UserAgent;
            webRequest.Method = "POST";
            byte[] bytes = Encoding.ASCII.GetBytes(parameters);
            Stream os = null;
            try
            {
                webRequest.ContentLength = bytes.Length;
                os = webRequest.GetRequestStream();
                os.Write(bytes, 0, bytes.Length);
            }
            catch (WebException ex)
            {
                // request error
            }
            finally
            {
                if (os != null)
                {
                    os.Close();
                }
            }

            try
            {
                WebResponse webResponse = webRequest.GetResponse();
                if (webResponse == null)
                { return null; }
                StreamReader sr = new StreamReader(webResponse.GetResponseStream(), Encoding.GetEncoding("ISO-8859-2"));
                return sr.ReadToEnd().Trim();
            }
            catch (WebException ex)
            {
                // response error
            }
            return null;
        }

        // zapis informacji o operacji do pliku
        private void WriteLog(string sData)
        {
            try
            {
                if (this.ImportLogFile != string.Empty & this.ImportLogFile != null)
                {
                    using (StreamWriter sw = File.AppendText(this.ImportLogFile))
                    {
                        sw.WriteLine("[" + DateTime.Now + "] " + sData);
                    }
                }
            }
            catch (Exception eW)
            {
                this.ImportException = eW;
                this.ImportError = "Nie można zapisać dziennika zdarzeń (najlepiej wyłącz go).";
                this.iStatus = ImporterStatus.ImportFailed;
                this.ImportFailed(null, new EventArgs());
            }
        }

        private BookAliases _ba;

        // przy uruchamianiu pluginu
        public void StartInitializing(BookAliases aliases)
        {
            this.lsTlumaczenia.Clear();
            this.lsKsiegi.Clear();

            this._ba = aliases;
            bool bPobranySkorowidz = true;
            string sHTML = string.Empty;
            try
            {
                WebClient wc = new WebClient { Encoding = Encoding.GetEncoding("ISO-8859-2") };
                wc.Headers.Add("user-agent", this.UserAgent);
                
                sHTML = wc.DownloadString("http://www.biblia.info.pl/menu.php");
            }
            catch (Exception)
            {
                bPobranySkorowidz = false;
            }

            if (bPobranySkorowidz)
            {
                // wycinamy zawartość tagów SELECT z nazwami tłumaczeń
                Match rM = Regex.Match(sHTML, @"<SELECT NAME=""tlumaczenie""(.*?)</SELECT>");
                while (rM.Success)
                {
                    sHTML = rM.ToString();
                    rM = rM.NextMatch();
                }

                // wycinamy i przepakowywujemy informacje o tłumaczeniach
                Match rM2 = Regex.Match(sHTML, @"<OPTION VALUE=""(.*?)"">(.*?)</OPTION>");
                List<string> sID = new List<string>();
                List<string> sNazwa = new List<string>();
                while (rM2.Success)
                {
                    if (rM2.Groups[1].Value != string.Empty)
                    {
                        sID.Add(rM2.Groups[1].Value);
                        sNazwa.Add(rM2.Groups[2].Value.Replace("(", "[").Replace(")", "]"));
                    }
                    rM2 = rM2.NextMatch();
                }

                // pakujemy tłumaczenia do listy dla hosta
                if (sID.Count == 0)
                {
                    this.ImportError = "Nie udało się poprawnie sparsować strony z tłumaczeniami.";
                    this.iStatus = ImporterStatus.InvalidData;
                    this.InitializingFailed(null, new EventArgs());
                }
                else
                {
                    this.iTlumaczenia = new ImporterAvailableTranslation[sID.Count];
                    for (byte i = 0; i < sID.Count; i++)
                    {
                        BibleLanguage blJezyk = new BibleLanguage("pl", "Polish", "Polski", "Rozdział {0}");
                        ImporterAvailableTranslation iatTlumaczenie = new ImporterAvailableTranslation(blJezyk, sID[i]);
                        iatTlumaczenie.Description = sNazwa[i];
                        iatTlumaczenie.UserFriendlyName = sNazwa[i];
                        iTlumaczenia[i] = iatTlumaczenie;

                        // tworzymy prywatną listę tłumaczeń na użytek późniejszy (przydatne przy imporcie konkretnego tłumaczenia)
                        this.lsTlumaczenia.Add(sID[i], sNazwa[i]);
                    }
                    this.iStatus = ImporterStatus.Ready;
                    this.InitializingComplete(null, new EventArgs());
                }
            }
            else
            {
                this.ImportError = "Nie można ustanowić połączenia ze stroną zawierającą tłumaczenia.";
                this.iStatus = ImporterStatus.NoConnection;
                this.InitializingFailed(null, new EventArgs());
            }
        }

        private bool bDoWorkStop = false;

        private void Import_DoWork(object sender, DoWorkEventArgs e)
        {
            BibleLanguage bl = new BibleLanguage("pl", "Polish", "Polski", "Rozdział {0}"); // niestety nie jestem w stanie stwierdzić języka
            BibleTranslation bt = new BibleTranslation(bl, "clash_importer_" + this.sImportowaneTlumaczenie); // nazwa pliku
            bt.Description = this.lsTlumaczenia[this.sImportowaneTlumaczenie] + " (zaimportowana ze strony http://biblia.info.pl)"; // opis przycisku wyświetlanego na pasku
            bt.FullName = this.lsTlumaczenia[this.sImportowaneTlumaczenie]; // nazwa biblii wyświetlana na pasku

            int iWszystkieRozdzialy = 0; 
            BackgroundWorker worker = sender as BackgroundWorker;
            for (int y = 0; y < this.lsKsiegi.Count; y++)
            {
                // tworzymy nową księgę
                BibleBook bb = new BibleBook(y + 1);
                string[] s = this.lsKsiegi[y].Split('|');
                for (byte x = 0; x < Convert.ToByte(s[1]); x++)
                {
                    // tworzymy nowy rozdział
                    BibleChapter bc = new BibleChapter(x + 1);

                    if (!this.bDoWorkStop)
                    {
                        this.ImportBook = this._ba.GetBookMainName(y+1) + ", rozdział " + (x + 1).ToString() + " z " + s[1];
                        WriteLog("Pobieram księgę od identyfikatorze [" + s[0] + "] (" + this.ImportBook + ")");
                        System.Threading.Thread.Sleep(0);
                        string sParametry = "tlumaczenie=" + this.sImportowaneTlumaczenie + "&nw=tak&ks=" + s[0] + "&rozdzial=" + (x + 1).ToString();
                        this.WriteLog("POST: http://www.biblia.info.pl/cgi-bin/biblia-nawigator.cgi&" + sParametry);
                        string sHTML = this.HttpPost("http://www.biblia.info.pl/cgi-bin/biblia-nawigator.cgi", sParametry).Replace("\r", "").Replace("\n", "").Replace("\t", "");

                        if (sHTML != null)
                        {
                            this.WriteLog(sHTML.Replace("\r", "").Replace("\n", "").Replace("\t", ""));

                            // sprawdzamy czy są wersety
                            if (!Regex.IsMatch(sHTML, "Błąd: "))
                            {
                                // parsowanie danych
                                string sWersetyPattern = @"SPAN class=""nrWersetu"">\((.*?)\)</SPAN> (.*?)<";
                                this.WriteLog("Parsuję wersety na podstawie paternu [" + sWersetyPattern + "]");
                                Match rM = Regex.Match(sHTML, sWersetyPattern);
                                int iWerset = 0;

                                while (rM.Success)
                                {
                                    if (rM.Groups[1].Value != string.Empty)
                                    {
                                        iWerset++;
                                        this.WriteLog("Znalazłem werset: [" + iWerset.ToString() + "] o treści: [" + rM.Groups[2].Value + "]");
                                        BibleVerse bv = new BibleVerse(iWerset, rM.Groups[2].Value);
                                        bc.AddVerse(bv);
                                    }
                                    rM = rM.NextMatch();
                                }
                            }
                            else
                            {
                                this.WriteLog("Nie znalazłem żadnych wersetów");
                                BibleVerse bv = new BibleVerse(1, "Tego rozdziału nie ma na stronie http://biblia.info.pl/biblia.php [" + DateTime.Now + "]");
                                bc.AddVerse(bv);
                                bv = new BibleVerse(2, this.ImporterName + " wersja: " + this.WersjaWtyczki);
                                bc.AddVerse(bv);
                            }

                            iWszystkieRozdzialy++; 
                            worker.ReportProgress((int)(((double)iWszystkieRozdzialy / (double)1189) * 100));
                        }
                        else
                        {
                            this.bDoWorkStop = true;
                            this.WriteLog("BŁĄD: Nie udało się pobrać strony (możliwe, że połączenie zostało zerwane)");
                            this.ImportError = "Nie udało się pobrać danych z serwera (zerwane połączenie?).";
                            this.iStatus = ImporterStatus.ImportFailed;
                            this.ImportFailed(null, new EventArgs());
                        }
                    }
                    bb.AddChapter(bc);
                }
                bt.AddBook(bb);
            }

            if (this.iStatus != ImporterStatus.ImportFailed)
            {
                this.WriteLog("Zapisuję biblię do pliku");
                this.ImportedBible = bt;
                WriteLog("-- STOP: " + this.ImporterName + " wersja: " + this.WersjaWtyczki);
                this.ImportComplete(null, new EventArgs());
            }
            else
                this.ImportFailed(null, new EventArgs());
        }

        private void Import_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.ImportProgressPercent = e.ProgressPercentage;
            this.ImportProgressChanged(null, new EventArgs());
        }

        private string sImportowaneTlumaczenie = null;

        private BackgroundWorker bw_worker;

        public void StartImport(string translationName)
        {
            this.bDoWorkStop = false;
            WriteLog("-- START: " + this.ImporterName + " wersja: " + this.WersjaWtyczki);
            this.iStatus = ImporterStatus.ImportInProgress;
            WriteLog("Rozpoczęty import tłumaczenia o identyfikatorze ["+translationName+"]");
            this.sImportowaneTlumaczenie = translationName;
            this.bw_worker = new BackgroundWorker();
            this.bw_worker.WorkerSupportsCancellation = true;
            this.bw_worker.WorkerReportsProgress = true;
            this.bw_worker.ProgressChanged += Import_ProgressChanged;
            this.bw_worker.DoWork += Import_DoWork;
            this.bw_worker.RunWorkerAsync();
        }

        public void AbortImport()
        {
            this.bDoWorkStop = true;
            this.bw_worker.CancelAsync();
            this.WriteLog("Operacja importu anulowana przez użytkownika");
            this.iStatus = ImporterStatus.ImportFailed;
            this.ImportError = "Operacja importu anulowana przez użytkownika";
            this.ImportFailed(null, new EventArgs());
        }

        public event EventHandler ImportComplete;
        public event EventHandler ImportFailed;
        public event EventHandler ImportProgressChanged;
        public event EventHandler InitializingComplete;
        public event EventHandler InitializingFailed;
    }
}
