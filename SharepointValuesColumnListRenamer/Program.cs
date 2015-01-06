using System;
using System.Collections.Generic;
using System.Net;
using SharepointValuesColumnListRenamer.Manager;
using SharepointValuesColumnListRenamer.Options;

namespace SharepointValuesColumnListRenamer
{
    class Program
    {
        // sample args => -w "http://spwebsite/_vti_bin/Lists.asmx" -d "DOMAIN" -u "USERNAME" -p "PASSWORD" -l "LISTNAME" -c "B_Email" -v "teamdebug@spwebsite.com" -s
        // sample args => -w "http://spwebsite/" -d "DOMAIN" -u "USERNAME" -p "PASSWORD" -l "News" -c "Title|Abstract" -v "UpdateWithConsole2|UpdateWithConsole3" -s -m
        static void Main(string[] args)
        {
            var options = new SPOptions();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                Resume(options);

                if (options.Start == false)
                {
                    Console.WriteLine("if ok, relaunch this cmd with arg : '-s'");
                }
                else
                {
                    Console.WriteLine("Launch Update ? (y/n)");
                    String response = Console.ReadLine();
                    if (response.ToLower() == "y" || response.ToLower() == "yes")
                    {
                        // WebSite URL
                        String url = AddPrefixAsmx(options.WebSite);
                        SetWebSiteURL(url);

                        if (CheckUrl(url))
                        {
                            SPReference.Lists client = new SPReference.Lists();

                            client.Credentials = new NetworkCredential
                            {
                                Domain = options.Domain,
                                UserName = options.UserName,
                                Password = options.Password
                            };

                            Dictionary<String, String> columnsAndValues = new Dictionary<string, string>();
                            if (options.Multi)
                            {
                                String[] sCs = options.Column.Split('|');
                                String[] sVs = options.Value.Split('|');
                                if (sCs.Length == sVs.Length)
                                {
                                    for (int i = 0; i < sCs.Length; i++)
                                    {
                                        columnsAndValues[sCs[i]] = sVs[i];
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Problem to parse Columns Values");
                                }
                            }
                            else
                            {
                                columnsAndValues[options.Column] = options.Value;
                            }

                            // Launch Update
                            SharepointManager.ListToUpdate(client, options.ListName, columnsAndValues);
                        }
                        else 
                        {
                            Console.WriteLine("Impossible URL access : " + url);
                        }
                    }
                    else
                    {
                        return;
                    }
                }
            }
            else if(args.Length == 0) 
            {
                SampleCmd();
            }
            else
            {
                Console.WriteLine("Problem to parse arguments or you missed one");
                SampleCmd();
            }

            Console.ReadLine();
        }
       
        static void Resume(SPOptions options)
        {
            Console.WriteLine("You set this Website URL : " + options.WebSite);
            Console.WriteLine("You set this domain : " + options.Domain );
            Console.WriteLine("You set this username : " + options.UserName);
            Console.WriteLine("You set this password : " + options.Password);
            Console.WriteLine("You set this listname : " + options.ListName);
            Console.WriteLine("You set this column : " + options.Column);
            Console.WriteLine("You set this value : " + options.Value);
            Console.WriteLine("You set multiple columns values to : " + options.Multi);
        }

        static void SampleCmd()
        {
            Console.WriteLine("Use arguments like this : -w WEBSITEURL -d DOMAIN -u USERNAME -p PASSWORD -l LISTNAME -c COLUMNNAME -v NEWVALUE");
        }

        // Update app.config with new Website service URL
        static void SetWebSiteURL(String url) {
            System.Xml.XmlDocument xml = new System.Xml.XmlDocument();
            xml.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            System.Xml.XmlNode node;
            node = xml.SelectSingleNode("configuration/applicationSettings/SharepointValuesColumnListRenamer.Properties.Settings/setting[@name='SharepointValuesColumnListRenamer_SPReference_Lists']");
            node.ChildNodes[0].InnerText = url;
            xml.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
        }

        static String AddPrefixAsmx(String url)
        { 
            if(url.Contains("asmx"))
            {
                return url;
            }
            else
            {
                // default list url
                url = url[url.Length - 1].Equals('/') ? url : url + "/";
                return url + "_vti_bin/Lists.asmx";
            }
        }

        static Boolean CheckUrl(String url)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "HEAD";
                
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                
                return (response.StatusCode == HttpStatusCode.OK);
            }
            catch 
            {
                return false;
            }
        }
    }
}
