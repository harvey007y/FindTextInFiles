﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IdealAutomate.Core {
  public class HotKeyRecord {
    public string[] HotKeys { get; set; }   
    public string Executable { get; set; }
    public string ExecuteContent { get; set; }
    public int ScriptID { get; set; }
  }
}
