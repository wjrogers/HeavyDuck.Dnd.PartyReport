using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;

namespace HeavyDuck.Dnd.PartyReport
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                if (DialogResult.OK == dialog.ShowDialog(this))
                {
                    foreach (var path in dialog.FileNames)
                        listBox1.Items.Add(path);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var items = new Dictionary<string, List<MagicItem>>();
            var output = @"C:\Temp\party.html";

            foreach (string path in listBox1.Items)
            {
                using (var fs = File.OpenRead(path))
                {
                    XPathDocument doc;
                    XPathNavigator nav;

                    doc = new XPathDocument(fs);
                    nav = doc.CreateNavigator();

                    var character_name = nav.SelectSingleNode("/D20Character/CharacterSheet/Details/name").Value.Trim();
                    var inventory = new List<MagicItem>();

                    for (var iter = nav.Select("//LootTally/loot[@count > 0]/RulesElement[@type = 'Magic Item']"); iter.MoveNext(); )
                    {
                        inventory.Add(new MagicItem
                        {
                            Name = iter.Current.SelectSingleNode("@name").Value.Trim(),
                            Url = iter.Current.SelectSingleNode("@url").Value.Trim(), // URLs in the XML are wrong
                            Level = iter.Current.SelectSingleNode("specific[@name = 'Level']").ValueAsInt,
                            MagicItemType = iter.Current.SelectSingleNode("specific[@name = 'Magic Item Type']").Value.Trim(),
                        });
                    }

                    items[character_name] = inventory;
                }
            }

            using (var fs = File.Open(output, FileMode.Create, FileAccess.Write))
            {
                using (var writer = XmlWriter.Create(fs))
                {
                    writer.WriteStartElement("html");

                    writer.WriteStartElement("head");
                    writer.WriteElementString("title", "West Chester D&D — Party Report");

                    writer.WriteStartElement("style");
                    writer.WriteAttributeString("type", "text/css");
                    writer.WriteValue(@"
body { font: 10pt Verdana; padding: 0; margin: 0; } 
div.character { float: left; margin: 1em; }
h1 { font-size: 1.2em; font-weight: bold; }
h2 { font-size: 1em; font-weight: bold; }");
                    writer.WriteEndElement(); // style

                    writer.WriteEndElement(); // head

                    writer.WriteStartElement("body");

                    foreach (var character in items.OrderBy(c => c.Key))
                    {
                        writer.WriteStartElement("div");
                        writer.WriteAttributeString("class", "character");

                        writer.WriteElementString("h1", character.Key);

                        var item_lookup = character.Value.ToLookup(g => g.MagicItemType);
                        var items_remaining = character.Value.ToList();

                        writer.WriteElementString("h2", "Weapons");
                        writer.WriteStartElement("ul");
                        foreach (var item in item_lookup["Weapon"].OrderByDescending(i => i.Level))
                        {
                            writer.WriteStartElement("li");
                            writer.WriteValue(string.Format("L{0:00} {1}", item.Level, item.Name));
                            writer.WriteEndElement(); // li

                            items_remaining.Remove(item);
                        }
                        writer.WriteEndElement(); // ul

                        writer.WriteElementString("h2", "Armor / Neck");
                        writer.WriteStartElement("ul");
                        foreach (var item in item_lookup["Armor"].Concat(item_lookup["Neck Slot Item"]).OrderByDescending(i => i.Level).OrderBy(i => i.MagicItemType))
                        {
                            writer.WriteStartElement("li");
                            writer.WriteValue(string.Format("L{0:00} {1}", item.Level, item.Name));
                            writer.WriteEndElement(); // li

                            items_remaining.Remove(item);
                        }
                        writer.WriteEndElement(); // ul

                        writer.WriteElementString("h2", "Other");
                        writer.WriteStartElement("ul");
                        foreach (var item in items_remaining.OrderBy(i => i.Name).OrderByDescending(i => i.Level))
                        {
                            writer.WriteStartElement("li");
                            writer.WriteValue(string.Format("L{0:00} {1}", item.Level, item.Name));
                            writer.WriteEndElement(); // li
                        }
                        writer.WriteEndElement(); // ul

                        writer.WriteEndElement(); //div
                    }

                    writer.WriteEndDocument();
                }
            }

            Process.Start(output);
        }

        private class MagicItem
        {
            public string Name { get; set; }
            public string Url { get; set; }
            public int Level { get; set; }
            public string MagicItemType { get; set; }
        }
    }
}
