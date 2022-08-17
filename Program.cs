// See https://aka.ms/new-console-template for more information
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Text.Json;
/**
* DISCLAIMERS
* This code WILL break the language and company's coding conventions
* for the sake of testing whenever the desired features can be implemented.
* Further verions of this code will polish the codebase, and production versions
* (if they ever come out) will follow all coventions and rules of the language
* and Veevart. So please, bear with me the following eye breaking prototype code.
*/
namespace App
{
    public class Response
    {
        public string CompanyLogo {get; set;}
        public string ViewGroup {get; set;}
        public Dictionary<String, Artwork> Artworks {get; set;}
    }

    public class Artwork {
        public string Title {get; set;}
        public string Artist {get; set;}
        public int Year {get; set;}
        public string Medium {get; set;}
        public string Movement {get; set;}
        public string Dimensions {get; set;}
        public string File {get; set;}
    }
    internal class Program
    {
        static void Main(string[] args)
        {
            FindReplaceOptions imageOptions = new FindReplaceOptions();
            imageOptions.ReplacingCallback = new TextToImage();
            //SECTION Parsing the data
            string text = System.IO.File.ReadAllText("./json/data.json");
            Response? res = new Response();
            
            res = JsonSerializer.Deserialize<Response>(text);

            Console.WriteLine($"Artworks send for view group ${res.ViewGroup}");
            foreach(KeyValuePair<string, Artwork> artwork in res.Artworks)
            {
                Console.WriteLine($"{artwork.Key}:");
                Console.WriteLine($"\t{artwork.Value.Title}");
                Console.WriteLine($"\t{artwork.Value.Artist}");
                Console.WriteLine($"\t{artwork.Value.Year}");
                Console.WriteLine($"\t{artwork.Value.Medium}");
                Console.WriteLine($"\t{artwork.Value.Movement}");
                Console.WriteLine($"\t{artwork.Value.Dimensions}");
            }
            //!SECTION

            //SECTION Parsing the document
            Document doc = new Document("templates/Print_List_Template_A.docx");
            Body body = doc.FirstSection.Body;
            // Parse the Headers and Footers first

            // Then the cover's title

            // Finally, parse the actual content
            Table table = (Table)body.GetChild(NodeType.Table, 0, false);
            Row row = table.FirstRow;
            for (int i = 1; i < res.Artworks.Count; i++)
            {
                table.AppendChild(row.Clone(true));
            }
            foreach (KeyValuePair<String, Artwork> ArtworkInfo in res.Artworks)
            {
                Cell leftCell = row.FirstCell;
                Cell rightCell = row.LastCell;
                Artwork Artwork = ArtworkInfo.Value;
                // Left Cell
                leftCell.Range.Replace("{!Artwork.Image}", $"images/artworks/{Artwork.File}", imageOptions);
                // Right Cell
                rightCell.Range.Replace("{!Artwork.Title}", Artwork.Title);
                rightCell.Range.Replace("{!Artwork.Artist}", Artwork.Artist);
                rightCell.Range.Replace("{!Artwork.Year}", Artwork.Year.ToString());
                rightCell.Range.Replace("{!Artwork.Medium}", Artwork.Medium);
                rightCell.Range.Replace("{!Artwork.Movement}", Artwork.Movement);
                rightCell.Range.Replace("{!Artwork.Dimentions}", Artwork.Dimensions);
                // Move to the next cell
                row = (Row)row.NextSibling;
            }

            doc.Save("outputs/PrintListE.pdf");
            //!SECTION
        }
    }

    public class TextToImage : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            Node currentNode = args.MatchNode;

            if (args.MatchOffset > 0)
                currentNode = splitRun((Run)currentNode, args.MatchOffset);

            DocumentBuilder builder = new DocumentBuilder(args.MatchNode.Document as Document);
            builder.MoveTo(currentNode);
            Shape img = builder.InsertImage(args.Replacement);
            img.VerticalAlignment = VerticalAlignment.Center;
            args.Replacement = "";
            return ReplaceAction.Replace;
        }

        static Run splitRun(Run run, int position)
        {
            Run otherRun = (Run)run.Clone(true);
            otherRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(otherRun, run); 
            return otherRun;
        }
    }
}
