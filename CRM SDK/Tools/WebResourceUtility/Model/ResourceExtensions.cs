using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WebResourceUtility.Model
{
    public static class ResourceExtensions
    {
        /// <summary>
        /// validFileExtensions : a utility array for valid file extensions representing Web Resources
        /// </summary>
        public static String[] ValidExtensions = { ".css", ".xml", ".gif", ".htm", ".html", ".ico", ".jpg", ".png", ".js", ".xap", ".xsl", ".xslt" };

         //Provides the integer value for the type of Web Resource
        public enum WebResourceType
        {
            Html = 1,
            Css = 2,
            JScript = 3,
            Xml = 4,
            Png = 5,
            Jpg = 6,
            Gif = 7,
            Silverlight = 8,
            Stylesheet_XSL = 9,
            Ico = 10
        }

        public static WebResourceType ConvertStringExtension(string extensionValue)
        {            
            switch (extensionValue.ToLower())
            {
                case "css":
                    return WebResourceType.Css;                    
                case "xml":
                    return WebResourceType.Xml;                    
                case "gif":
                    return WebResourceType.Gif;                    
                case "htm":
                    return WebResourceType.Html;                    
                case "html":
                    return WebResourceType.Html;                    
                case "ico":
                    return WebResourceType.Ico;                  
                case "jpg":
                    return WebResourceType.Jpg;                    
                case "png":
                    return WebResourceType.Png;                   
                case "js":
                    return WebResourceType.JScript;                    
                case "xap":
                    return WebResourceType.Silverlight;                   
                case "xsl":
                    return WebResourceType.Stylesheet_XSL;
                case "xslt":
                    return WebResourceType.Stylesheet_XSL;
                default:
                    throw new ArgumentOutOfRangeException(string.Format("\"{0}\" is not recognized as a valid file extension for a Web Resource.", extensionValue.ToLower()));

            }
        }
    }
}
