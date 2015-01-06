using System;
using System.Collections.Generic;
using System.Web.Services.Protocols;
using System.Xml;

namespace SharepointValuesColumnListRenamer.Manager
{
    static class SharepointManager
    {
        /// <summary>
        /// Update List
        /// </summary>
        /// <param name="client"></param>
        /// <param name="listName"></param>
        /// <param name="columnsAndValues"></param>
        public static void ListToUpdate(SPReference.Lists client, String listName, Dictionary<String, String> columnsAndValues)
        {
            //
            XmlNode listItems = null;
            XmlDocument xmlDoc = new System.Xml.XmlDocument();
            XmlNode queryWhere = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");

            try
            {
                queryWhere.InnerXml = SharepointManager.GenerateQueryWhere(columnsAndValues);
            }
            catch (XmlException xmlEx)
            {
                Console.WriteLine("Xml Format Error Query Where ( " + xmlEx.Message + " )");
            }

            listItems = SharepointManager.GetListItemsWeb(client, listName, queryWhere);
            // First time
            SharepointManager.UpdateList(client, listItems, queryWhere, listName, columnsAndValues);
            // loop
            listItems = SharepointManager.GetListItemsWeb(client, listName, queryWhere);
            while (listItems != null && listItems.ChildNodes[1].Attributes["ItemCount"].Value != "0")
            {
                SharepointManager.UpdateList(client, listItems, queryWhere, listName, columnsAndValues);
                listItems = SharepointManager.GetListItemsWeb(client, listName, queryWhere);
            }
            Console.WriteLine("Ending");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="client"></param>
        /// <param name="listName"></param>
        /// <param name="queryWhere"></param>
        /// <returns></returns>
        private static XmlNode GetListItemsWeb(SPReference.Lists client, String listName, XmlNode queryWhere)
        {
            XmlNode listItems = null;

            // Check if connection is
            try
            {
                // Get
                listItems = client.GetListItems(listName, null, queryWhere, null, "100", null, null);
            }
            catch (SoapException soapEx)
            {
                Console.WriteLine("Soap Ex ( " + soapEx.Message + " )");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error ( " + ex.Message + " )");
            }

            return listItems;
        }

        private static void UpdateList(SPReference.Lists client, XmlNode listItems, XmlNode queryWhere, String listName, Dictionary<String, String> columnsAndValues)
        {
            if (listItems != null && listItems.ChildNodes[1].Attributes["ItemCount"].Value != "0")
            {
                // Prepare Moderation
                XmlDocument docMod1 = new XmlDocument();
                XmlElement batchElementMod1 = docMod1.CreateElement("Batch");

                // Prepare Update
                XmlDocument docUp = new XmlDocument();
                XmlElement batchElementUp = docMod1.CreateElement("Batch");

                // Prepare Moderation 2
                XmlDocument docMod2 = new XmlDocument();
                XmlElement batchElementMod2 = docMod1.CreateElement("Batch");

                String innerXmlModeration = "",
                       innerXmlModeration2 = "",
                       innerXmlUpdate = "";

                foreach (XmlNode node in listItems)
                {
                    if (node.Name == "rs:data")
                    {
                        for (int f = 0; f < node.ChildNodes.Count; f++)
                        {
                            if (node.ChildNodes[f].Name == "z:row")
                            {
                                // Change Moderation Status
                                innerXmlModeration += SharepointManager.GenerateInnerXmlModeration(node.ChildNodes[f], true);
                                innerXmlModeration2 += SharepointManager.GenerateInnerXmlModeration(node.ChildNodes[f], false);
                                // Update Values
                                innerXmlUpdate += SharepointManager.GenerateInnerXmlUpdate(node.ChildNodes[f], columnsAndValues);
                            }
                        }
                    }
                }

                #region Moderation

                try
                {
                    batchElementMod1.InnerXml = innerXmlModeration;
                }
                catch (XmlException xmlEx)
                {
                    Console.WriteLine("Xml Format Error Moderation ( " + xmlEx.Message + " )");
                }

                // Exec Update
                try
                {
                    client.UpdateListItems(listName, batchElementMod1);

                    Console.WriteLine("Update Moderation Completed");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error Update Uncompleted Moderation ( " + e.Message + " )");
                }

                #endregion

                #region Update

                try
                {

                    batchElementUp.InnerXml = innerXmlUpdate;
                }
                catch (XmlException xmlEx)
                {
                    Console.WriteLine("Xml Format Error Update ( " + xmlEx.Message + " )");
                }

                // Exec Update
                try
                {
                    client.UpdateListItems(listName, batchElementUp);

                    Console.WriteLine("Update Update Completed");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error Update Uncompleted Update( " + e.Message + " )");
                }

                #endregion

                #region Moderation2

                try
                {

                    batchElementMod2.InnerXml = innerXmlModeration2;
                }
                catch (XmlException xmlEx)
                {
                    Console.WriteLine("Xml Format Error Moderation2 ( " + xmlEx.Message + " )");
                }

                // Exec Update
                try
                {
                    client.UpdateListItems(listName, batchElementMod2);

                    Console.WriteLine("Update Moderation 2 Completed");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error Update Uncompleted Moderation2 ( " + e.Message + " )");
                }

                #endregion
            }
            else
            {
                Console.WriteLine("No row updated !");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="node"></param>
        /// <param name="first"></param>
        /// <returns></returns>
        private static String GenerateInnerXmlModeration(XmlNode node, Boolean first)
        {
            String innerXml = "<Method ID='UpdateModerationWithSharepointManager_";

            if (first == false)
            {
                innerXml = "<Method ID='UpdateModeration2WithSharepointManager_";
            }

            innerXml += node.Attributes["ows_ID"].Value + "' Cmd='Moderate'>" +
                        "<Field Name='ID'>" + node.Attributes["ows_ID"].Value + "</Field>";


            innerXml += "<Field Name='_ModerationStatus'>0</Field>";

            innerXml += "</Method>";

            return innerXml;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="node"></param>
        /// <param name="columnsAndValues"></param>
        /// <returns></returns>
        private static String GenerateInnerXmlUpdate(XmlNode node, Dictionary<String, String> columnsAndValues)
        {
            String innerXml = "<Method ID='UpdateWithSharepointManager_" + node.Attributes["ows_ID"].Value + "' Cmd='Update'>" +
                                    "<Field Name='ID'>" + node.Attributes["ows_ID"].Value + "</Field>";

            // Dictionnary
            foreach (var kvalue in columnsAndValues)
            {
                innerXml += "<Field Name='" + kvalue.Key + "'>" + kvalue.Value + "</Field>";
            }

            innerXml += "</Method>";

            return innerXml;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="columnsAndValues"></param>
        /// <returns></returns>
        private static String GenerateQueryWhere(Dictionary<String, String> columnsAndValues)
        {
            String queryWhere = "<Where><Or>";

            queryWhere += "<Neq><FieldRef Name='_ModerationStatus' /><Value Type='Number'>0</Value></Neq>";

            // Dictionnary
            foreach (var kvalue in columnsAndValues)
            {
                String type = "Text";
                try
                {
                    int.Parse(kvalue.Value);
                    type = "Number";
                }
                catch
                {
                    type = "Text";
                }
                queryWhere += "<Neq><FieldRef Name='" + kvalue.Key + "' /><Value Type='" + type + "'>" + kvalue.Value + "</Value></Neq>";
            }

            queryWhere += "</Or></Where>";

            return queryWhere;
        }
    }
}
