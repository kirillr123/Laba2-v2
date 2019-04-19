using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Paging
{
    [Serializable]
    public class Note
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public string Object { get; set; }
        public bool IsConfidential { get; set; }
        public bool IsInteger { get; set; }
        public bool IsAccesible { get; set; }

        public bool Equals(Note s)
        {
            if(!Id.Equals(s.Id))
            {
                return false;
            }
            if (!Name.Equals(s.Name))
            {
                return false;
            }
            if (!Description.Equals(s.Description))
            {
                return false;
            }
            if (!Source.Equals(s.Source))
            {
                return false;
            }
            if (!Object.Equals(s.Object))
            {
                return false;
            }
            if (IsConfidential!=s.IsConfidential)
            {
                return false;
            }
            if (IsInteger!=s.IsInteger)
            {
                return false;
            }
            if (IsAccesible!=s.IsAccesible)
            {
                return false;
            }
            return true;
        }
    }
}
