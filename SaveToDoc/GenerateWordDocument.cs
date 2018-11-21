using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace SaveToDoc
{
    public class GenerateWordDocument
    {
        public static object ReadOnly = false;
        public static object Convert = false;
        public static object IsVisible = false;
        public static object FileFormat = WdSaveFormat.wdFormatPDF;
        public static object SaveChangesToTemplate = false;

        /// <summary>
        /// Generates a word doc from a given tamplate and replaces the tags with it's given value
        /// </summary>
        /// <param name="templateLocation"></param>
        /// <param name="newFileNameWithLocation"></param>
        /// <param name="tags"></param>
        public static void GenerateWordDoc(object templateLocation, object newFileNameWithLocation, List<ReplaceTag> tags, bool readOnly)
        {
            object missing = Missing.Value;
            object readOnlySave = readOnly;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = null;

            wordApp.Visible = (bool)IsVisible;

            wordDoc = wordApp.Documents.Open(
                        templateLocation,
                        ref missing,
                        ref ReadOnly,
                        ref missing,
                        ref missing,
                        ref missing,
                        ref missing,
                        ref missing,
                        ref missing,
                        ref missing,
                        ref missing,
                        ref IsVisible,
                        ref missing,
                        ref missing,
                        ref missing
                    );

            wordDoc.Activate();

            foreach (ReplaceTag replaceTag in tags)
            {
                FindAndReplace(wordApp, replaceTag.Replace, replaceTag.Value);
            }

            wordDoc.SaveAs2(
                ref newFileNameWithLocation,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref readOnlySave,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing
            );

            wordApp.Quit(
                ref SaveChangesToTemplate,
                ref missing,
                ref missing
            );
        }

        private static void FindAndReplace(Microsoft.Office.Interop.Word.Application word, object find, object replace)
        {
            object _match_case = true;
            object _match_whole_word = true;
            object _match_wild_cards = false;
            object _match_sound_like = false;
            object _match_all_forms = false;
            object _forward = true;
            object _format = false;
            object _match_kashida = false;
            object _match_diactitics = false;
            object _match_alef_hamza = false;
            object _matchcontrol = false;
            object _read_only = false;
            object _visible = true;
            object _replace = 2;
            object _wrap = 1;

            word.Selection.Find.Execute(
                ref find,
                ref _match_case,
                ref _match_whole_word,
                ref _match_wild_cards,
                ref _match_sound_like,
                ref _match_all_forms,
                ref _forward,
                ref _wrap,
                ref _format,
                ref replace,
                ref _replace,
                ref _match_kashida,
                ref _match_diactitics,
                ref _match_alef_hamza,
                ref _matchcontrol
            );
        }

        /// <summary>
        /// Combines multiple documents into one with page brake
        /// </summary>
        /// <param name="fileNamesLocation"></param>
        /// <param name="newFileNameWithLocation"></param>
        public static void CombineFiles(List<string> fileNamesLocation, object newFileNameWithLocation, bool readOnly)
        {
            object missing = Missing.Value;
            string fileLocation = fileNamesLocation[0];
            object readOnlySave = readOnly;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = null;
            Microsoft.Office.Interop.Word.Document wordDocCopy = null;
            
            wordApp.Visible = (bool)IsVisible;

            wordDoc = wordApp.Documents.Add();

            wordDoc.Activate();

            for (int i = 0; i < fileNamesLocation.Count; i++)
            {
                string copyFileLocation = fileNamesLocation[i];

                wordDocCopy = wordApp.Documents.Open(
                            copyFileLocation,
                            ref Convert,
                            ref readOnlySave,
                            ref missing,
                            ref missing,
                            ref missing,
                            ref missing,
                            ref missing,
                            ref missing,
                            ref missing,
                            ref missing,
                            ref IsVisible,
                            ref missing,
                            ref missing,
                            ref missing
                        );

                Microsoft.Office.Interop.Word.Range range = wordDocCopy.Content;
                range.Copy();

                wordDoc.Range(wordDoc.Content.End - 1).Paste();
                wordDoc.Range(wordDoc.Content.End - 1).InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

                wordDocCopy.Close();
            }

            wordDoc.SaveAs2(
                ref newFileNameWithLocation,
                ref FileFormat,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref readOnlySave,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing
            );

            wordApp.Quit(
                ref SaveChangesToTemplate,
                ref missing,
                ref missing
            );
        }
    }

    public class ReplaceTag
    {
        public ReplaceTag(string replace, object value)
        {
            Replace = replace;
            Value = value;
        }

        public string Replace { get; set; }
        public object Value { get; set; }
    }
}