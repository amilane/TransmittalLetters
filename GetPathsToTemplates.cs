using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransmitLetter
{
  class GetPathsToTemplates
  {
    public List<string> getPathsToTemplates()
    {
      var paths = File.ReadAllLines("PathsToTemplates.txt").ToList();
      return paths;
    }
  }
}
