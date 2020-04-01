using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using System.Globalization;

namespace Microsoft.Crm.Sdk.Samples
{
    public class RegexValidationRule : ValidationRule
    {
        private string _pattern;
        private Regex _regex;
        private string _message;

        public string Pattern
        {
            get { return _pattern; }
            set
            {
                _pattern = value;
                _regex = new Regex(_pattern,
                    (RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase));
                
            }
        }

        public string Message
        {
            get { return _message; }
            set
            {
                _message = value;
            }
        }
        

        public override ValidationResult Validate(object value, CultureInfo ultureInfo)
        {
            if (value != null)
            {
                if (_regex.IsMatch(value.ToString()))
                {
                    return new ValidationResult(false, _message);
                }
            }
            
            return new ValidationResult(true, null);
            
        }
    }

}
