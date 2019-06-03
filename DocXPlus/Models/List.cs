using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace DocXPlus
{
    public class List
    {
        private readonly DocX _document;

        private int NumberingLevelReference = 0;
        private int NumberingId = 1;

        private Paragraph AddedParagrah;

        public List(DocX document)
        {
            _document = document;
        }

        public List AppendItems(params string[] items)
        {
            foreach (var item in items)
            {
                AppendItem(item);
            }

            return this;
        }

        public List AppendItems(params (string Item, Action<Paragraph> ConfigParagraph)[] items)
        {
            foreach (var item in items)
            {
                AppendItem(item.Item);
                item.ConfigParagraph(AddedParagrah);
            }

            return this;
        }

        public List AppendBoldItems(params string[] items)
        {
            foreach (var item in items)
            {
                AppendItem(item);
                AddedParagrah.Bold();
            }

            return this;
        }

        public List AppendItem(string item)
        {
            var paragraph = _document.AddParagraph(item);
            AddedParagrah = paragraph;

            var numberingProperties = paragraph.GetParagraphProperties().GetOrCreate<NumberingProperties>(true);

            numberingProperties.NumberingLevelReference = new NumberingLevelReference() { Val = NumberingLevelReference };
            numberingProperties.NumberingId = new NumberingId() { Val = NumberingId };

            return this;
        }

        public List AppendItem(string item, Action<List> addSubItems)
        {
            var paragraph = _document.AddParagraph(item);
            AddedParagrah = paragraph;

            var numberingProperties = paragraph.GetParagraphProperties().GetOrCreate<NumberingProperties>(true);

            numberingProperties.NumberingLevelReference = new NumberingLevelReference() { Val = NumberingLevelReference };
            numberingProperties.NumberingId = new NumberingId() { Val = NumberingId };

            var subList = new List(_document)
            {
                NumberingLevelReference = this.NumberingLevelReference + 1,
                NumberingId = this.NumberingId
            };

            addSubItems(subList);

            return this;
        }

        public List Bold()
        {
            AddedParagrah.Bold();
            return this;
        }

        public List Italic()
        {
            AddedParagrah.Italic();
            return this;
        }

        public List Underline(UnderlineType type)
        {
            AddedParagrah.Underline(type);
            return this;
        }
    }
}
