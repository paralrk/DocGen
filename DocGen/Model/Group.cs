using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocGen.Utils;


namespace DocGen.Model
{
    class Group : IEnumerable
    {
        public List<Components> CList { get; set; } // хранятся компоненты одного типа
        public string DesignatorType { get; set; } // C, DA, R 
        public string TypeDescription { get; set; } // Конденсатор, Микросхема, Резистор
        public string TypeDescriptions { get; set; } // Конденсаторы, Микросхемы, Резисторы

        public Dictionary<string, int> Manufactures { get; } = new Dictionary<string, int>();
        public List<GroupHeader> Headers { get; } = new List<GroupHeader>();

        // constructors
        Group(String designatorType, String typeDescription, String typeDescriptions) : base()
        {
            this.DesignatorType = designatorType;
            this.TypeDescription = typeDescription;
            this.TypeDescriptions = typeDescriptions;
        }

        public Group()
        {
            this.CList = new List<Components>();
            DesignatorType = "-";
            TypeDescription = "";
            TypeDescription = "Прочие";
        }

        public List<Components> GetComponents()
        {
            return CList;
        }
        public void Add(Components c)
        {
            CList.Add(c);
        }

        public bool IsEmpty()
        {
            if (CList.Count == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void Clear()
        {
            CList.Clear();
        } // clear

        public IEnumerator GetEnumerator()
        {
            return CList.GetEnumerator();
        }

        public void CountManufactures(int limit)
        {
            GroupHeader header = null;
            foreach (Components c in CList)
            {
                if (c.Part.Manufacturer != null)
                {
                    header = FindHeader(c);
                    if (header != null)
                    {
                        header.Count += 1;
                    }
                    else
                    {
                        header = new GroupHeader();
                        header.TypeDescription = c.Part.Type;
                        header.Manufacturer = c.Part.Manufacturer;
                        header.NoteNumber = c.NoteNumber;
                        header.NoteNumber1 = c.NoteNumber1;
                        header.Count = 1;
                        Headers.Add(header);
                    }
                }
            }

            for (int i = Headers.Count - 1; i >= 0; i--)
            {
                if (Headers[i].Count < limit)
                {
                    Headers.RemoveAt(i);
                }
            }
        }

        public bool IsInHeaders(Components c)
        {
            if (c == null)
            {
                return false;
            }
            if (FindHeader(c) != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private GroupHeader FindHeader(Components c)
        {
            return Headers.Find(
                h => h.TypeDescription.Equals(c.Part.Type)
                && h.Manufacturer.Equals(c.Part.Manufacturer)
                );
        }
        public GroupHeader GetHeader(Components c)
        {
            return FindHeader(c);
        }

    }
}
