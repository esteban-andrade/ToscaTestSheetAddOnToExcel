using System.Collections.Generic;
using System.Linq;
using Tricentis.TCAPIObjects.Objects;

namespace TestSheetAddOn
{
    public class TDElementWrapper
    {
        public TDStructElement TDElement { get; set; }
        public TDElementWrapper ParentWrapper { get; set; }

        public string Path { get { return string.Join(".", GetParentsAndMe()); } }

        private IEnumerable<string> GetParentsAndMe()
        {
            string[] myName = new string[] { TDElement.DisplayedName };
            if (ParentWrapper == null)
            {
                return myName;
            }
            return ParentWrapper.GetParentsAndMe().Concat(myName);
        }
    }

}
