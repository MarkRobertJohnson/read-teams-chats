using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;

namespace MarkTeamsChatsAsRead
{
    internal class CommandLineOptions
    {
        [Option('a', "all", Required = false, HelpText = "Force read all chats", Default = false)]
        public bool All { get; set; }

        [Option('p', "parallelism", Default = 1, Required = false, HelpText = "Specify how many thread to use when marking chats as read.  Default is 1.  NOTE: Too many threads can cause API throttling.")]
        public int Parellelism { get;set; }
    }
}
