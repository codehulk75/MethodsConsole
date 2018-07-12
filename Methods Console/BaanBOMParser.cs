﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Text.RegularExpressions;

namespace Methods_Console
{
    class BaanBOMParser : FileParser
    {
        public override string FileType { get; set; }
        public override string FileName { get; set; }
        public override string FullFilePath { get; set; }
        public string AssemblyName { get; private set; }
        public string DateOfListing { get; private set; }
        public string AssyDescription { get; private set; }
        public string Rev { get; private set; }
        public List<string> RouteList { get; private set; }
        public bool IsValid { get; private set; }
        public Dictionary<string, List<string>> BomMap { get; private set; } //Key = '<Findnum>:<Sequence>', List = [0]Part Number, [1]Operation, [2]AssyDescription, [3]comma-separated ref des's 
        public List<string> BomLines { get; private set; }

        public BaanBOMParser(string path)
        {
            AssemblyName = null;
            DateOfListing = null;
            AssyDescription = null;
            Rev = null;
            FullFilePath = path;
            FileName = Path.GetFileName(FullFilePath);
            FileType = Path.GetExtension(FullFilePath).ToLower();
            BomMap = new Dictionary<string, List<string>>();
            BomLines = new List<string>();
            RouteList = new List<string>();
            LoadBaanBom();
            ParseBaanBom();
            SetValid();         
        }

        private void SetValid()
        {
            if (!string.IsNullOrEmpty(AssemblyName) && !string.IsNullOrEmpty(AssyDescription) && !string.IsNullOrEmpty(DateOfListing) && !string.IsNullOrEmpty(Rev) && BomMap.Count > 0)
                IsValid = true;
            else
                IsValid = false;
        }

        private void ClearData()
        {
            FileType = null;
            FileName = null;
            FullFilePath = null;
            AssemblyName = null;
            DateOfListing = null;
            AssyDescription = null;
            Rev = null;
            RouteList.Clear();
            BomMap.Clear();
            BomLines.Clear();
        }

        private void LoadBaanBom()
        {
            string line;
            try
            {
                using (StreamReader sr = new StreamReader(FullFilePath))
                {
                    while ((line = sr.ReadLine()) != null)
                    {
                        BomLines.Add(line);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error creating BOM File Parser object.\nMake sure it's a valid BaaN BOM and try again.\nLoadBaanBom()\n" + e.Message, "BaanBOMParser.LoadBaanBOM()", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
            }
        }

        private void ParseBaanBom()
        {
            Regex reItemInfoRow= new Regex(@"1\s+\|\s*([0-9]{1,4})/\s+([0-9])\|(\S+)\s*\|(.*)\|Purchased\s+\|\s+([0-9]{1,4}).*\|.*\|.*\|.*\|.*\|.*\|.*\|\s*(\d+\.\d+)"); //Captures 1-FindNum, 2-Seq, 3-PN, 4-Operation, 5-BOMQty
            Regex reDate = new Regex(@"Date\s+:\s(\d{2}-\d{2}-\d{2})");
            Regex reRefDes = new Regex(@"^\s+\|\s(\w+)\s+\|\s+\|\s+\d{1,4}\.\d{4}\s+\|\r?$");
            Regex reAssemblyName = new Regex(@"Manufactured Item\s+:\s+(\S+)");
            Regex reRev = new Regex(@"Revision\s+:\s+(\S+)");
            Regex reRouteList = new Regex(@"^\s+\|\s+(\d{1,4})/\s+\d\s+\|\s+(\d{1,4})\s+\|(.*)\|\d{2}-\d{2}-\d{2}\s+\|\s+\|\s+\w+\s+\|\s+\|\s+\d{1,3}\s+\|\s+\d+\.\d+\s+\|\s+\d+\.\d+\s*\|\s+\|\r?$");
            Regex reDescription = new Regex(@"^Description\s+:\s+(.*)");
            bool bDateFound = false;
            bool bAssemblyNameFound = false;
            bool bRevFound = false;
            bool bDescriptionFound = false;
            string strCurrentFnSeq = null;

            try
            {
                foreach (string line in BomLines)
                {
                    if (bDateFound == false)
                    {
                        Match m = reDate.Match(line);
                        if (m.Success)
                        {
                            DateOfListing = m.Groups[1].Value;
                            bDateFound = true;
                            continue;
                        }
                    }
                    if (bAssemblyNameFound == false)
                    {
                        Match m = reAssemblyName.Match(line);
                        if (m.Success)
                        {
                            AssemblyName = m.Groups[1].Value;
                            bAssemblyNameFound = true;
                            continue;
                        }                          
                    }
                    if (bRevFound == false)
                    {
                        Match m = reRev.Match(line);
                        if (m.Success)
                        {
                            Rev = m.Groups[1].Value;
                            bRevFound = true;
                            continue;
                        }
                    }
                    if (bDescriptionFound == false)
                    {
                        Match m = reDescription.Match(line);
                        if (m.Success)
                        {
                            AssyDescription = m.Groups[1].Value.TrimEnd();
                            bDescriptionFound = true;
                            continue;
                        }
                    }
                    Match match = reItemInfoRow.Match(line);
                    if (match.Success)
                    {   
                        string strFnSeq = match.Groups[1].Value + ":" + match.Groups[2].Value;
                        string strPN = match.Groups[3].Value;
                        string strDesc = match.Groups[4].Value;
                        string strOp = match.Groups[5].Value;
                        string strBOMQty = match.Groups[6].Value.Trim();
                        strBOMQty = strBOMQty.Trim('0');
                        if (strBOMQty.EndsWith("."))
                            strBOMQty = strBOMQty.Trim('.');
                        if (strBOMQty.Equals(""))
                            strBOMQty = "0";
                        if (BomMap.ContainsKey(strFnSeq))
                            MessageBox.Show(string.Format("Duplicate Findnum/Sequence combination found!\nThis should not happen!  Check BOM:\n"
                                + "Key: = {0}, PN: = {4}\nExisting Values: {1}, {2}, {3}" + "\nNot adding new part to BOM",
                                strFnSeq, BomMap[strFnSeq][0], BomMap[strFnSeq][1], BomMap[strFnSeq][2], strPN), "BOM Error", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        else
                        {
                            BomMap.Add(strFnSeq, new List<string> { strPN, strOp, strDesc.Trim(), "", strBOMQty });
                            if (!string.IsNullOrEmpty(strCurrentFnSeq) && BomMap[strCurrentFnSeq].Count > 2)
                                BomMap[strCurrentFnSeq][3] = BomMap[strCurrentFnSeq][3].TrimEnd(',');

                            strCurrentFnSeq = strFnSeq;                        
                        }                           
                        continue;
                    }
                    Match matchRds = reRefDes.Match(line);
                    if (matchRds.Success)
                    {
                        BomMap[strCurrentFnSeq][3] += matchRds.Groups[1].Value + ",";
                        continue;
                    }
                    Match matchRouteList = reRouteList.Match(line);
                    if (matchRouteList.Success)
                    {
                        string strRouteStep = matchRouteList.Groups[1].Value + ":" + matchRouteList.Groups[2].Value + ":" + matchRouteList.Groups[3].Value.Trim();
                        RouteList.Add(strRouteStep);
                    }

                }
                BomMap.ElementAt(BomMap.Count - 1).Value[3] = BomMap.ElementAt(BomMap.Count - 1).Value[3].TrimEnd(',');
                foreach(string key in BomMap.Keys)
                {
                    if (BomMap[key][3].Equals(""))
                    {
                        BomMap[key][3] = "No Ref Des";
                    }
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Error parsing BAAN file.\nMake sure it's a valid BaaN BOM and try again.\nParseBaanBom()\n" + e.Message, strCurrentFnSeq, MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
                throw;
            }
        }
    }
}
