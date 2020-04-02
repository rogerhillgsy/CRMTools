//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Media;

namespace ModernOdataApp
{
  public class MainPageItem
  {
    public string Name { get; set; }

    //Return the Item name
    public string ItemName
    {
      get
      {
        return string.Format("{0:00}", Name);
      }
    }

    //Get the Item Image
    public string Image
    {
      get
      {
        return string.Format("ms-appx:///Assets/MainPage/{0}.png", Name);
      }
    }

    //Create the Item
    public UIElement XAMLItem
    {
      get
      {
        //XAML Item Uri...
        string xamlPath = string.Format("ms-appx:///Assets/MainPage/{0}.xaml", Name);
        Uri xamlUri = new Uri(xamlPath, UriKind.Absolute);

        //Create new canvas element for the rectangle and path to be loaded into...
        Canvas targetCanvas = new Canvas();

        //Load the contents of the XAML Card's .xaml file into the in memory canvas using App.LoadComponent
        App.LoadComponent(targetCanvas, xamlUri);

        //Wrap the new canvas up in a Viewbox so that it scales to fit the container.  
        Viewbox viewBox = new Viewbox();
        viewBox.Child = targetCanvas;

        //Return the Viewbox...
        return viewBox;
      }
    }

  }
}
