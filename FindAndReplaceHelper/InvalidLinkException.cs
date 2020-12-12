using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindAndReplaceHelper
{
    public class InvalidLinkException : Exception
    {
        public InvalidLinkException()
        {

        }

        public InvalidLinkException(string link) : 
            base(
                $"\nSorry {link} is not a valid link,\n please enter a valid SEEK job description link\n" +
                        "e.g 'https://www.seek.com.au/job/51117822?type=standard#searchRequestToken=6a1e962b-4904-4003-837b-2645622711a9'\n" +
                        "otherwise enter 'exit' to exit application\n\n"
                )
        {
        }
    }
}
