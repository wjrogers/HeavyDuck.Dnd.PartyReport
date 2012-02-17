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
                if (DialogResult.OK == dialog.ShowDialog(this))
                    listBox1.Items.Add(dialog.FileName);
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
                            Url = iter.Current.SelectSingleNode("@url").Value.Trim(),
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
                    writer.WriteStartElement("body");

                    foreach (var character in items)
                    {
                        writer.WriteStartElement("div");
                        writer.WriteAttributeString("style", "float: left;");

                        writer.WriteElementString("h1", character.Key);

                        var item_lookup = character.Value.ToLookup(g => g.MagicItemType);
                        var items_remaining = character.Value.ToList();

                        writer.WriteElementString("h2", "Weapons");
                        writer.WriteStartElement("ul");
                        foreach (var item in item_lookup["Weapon"].OrderBy(i => i.Name))
                        {
                            writer.WriteStartElement("li");
                            writer.WriteStartElement("a");
                            writer.WriteAttributeString("href", item.Url);
                            writer.WriteValue(string.Format("L{0:00} {1}", item.Level, item.Name));
                            writer.WriteEndElement(); // a
                            writer.WriteEndElement(); // li

                            items_remaining.Remove(item);
                        }
                        writer.WriteEndElement(); // ul

                        writer.WriteElementString("h2", "Armor");
                        writer.WriteStartElement("ul");
                        foreach (var item in item_lookup["Armor"].Concat(item_lookup["Neck Slot Item"]).OrderBy(i => Tuple.Create(i.MagicItemType, i.Name)))
                        {
                            writer.WriteStartElement("li");
                            writer.WriteStartElement("a");
                            writer.WriteAttributeString("href", item.Url);
                            writer.WriteValue(string.Format("L{0:00} {1}", item.Level, item.Name));
                            writer.WriteEndElement(); // a
                            writer.WriteEndElement(); // li

                            items_remaining.Remove(item);
                        }
                        writer.WriteEndElement(); // ul

                        writer.WriteElementString("h2", "Other");
                        writer.WriteStartElement("ul");
                        foreach (var item in items_remaining.OrderByDescending(i => i.Level))
                        {
                            writer.WriteStartElement("li");
                            writer.WriteStartElement("a");
                            writer.WriteAttributeString("href", item.Url);
                            writer.WriteValue(string.Format("L{0:00} {1}", item.Level, item.Name));
                            writer.WriteEndElement(); // a
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
