using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Markup;
using System.Windows.Data;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Globalization;

namespace Microsoft.Crm.Sdk.Samples
{
    [ValueConversion(typeof(ReadOnlyObservableCollection<ValidationError>), typeof(string))]
    public class ValidationErrorsToStringConverter : MarkupExtension, IValueConverter
    {
        public ValidationErrorsToStringConverter()
        {

        }
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return new ValidationErrorsToStringConverter();
        }
 
        public object Convert(object value, Type targetType, object parameter, 
            CultureInfo culture)
        {
            ReadOnlyObservableCollection<ValidationError> errors =
                value as ReadOnlyObservableCollection<ValidationError>;
 
            if (errors == null)
            {
                return string.Empty;
            }
 
            return string.Join("\n", (from e in errors 
                                      select e.ErrorContent as string).ToArray());
        }
 
        public object ConvertBack(object value, Type targetType, object parameter, 
            CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
