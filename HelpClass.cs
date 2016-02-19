//-----------------------------------------------------------------------
// <copyright file="HelpClass.cs" company="N/A">
//     Copyright (c) 2016 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

namespace VehicleInformationLookupTool
{
    using System.Windows.Forms;

    /// <summary>
    ///  Wrapper for the Windows Forms Help class
    /// </summary>
    public static class HelpClass
    {
        /// <summary>
        /// Open the help file to the specified topic
        /// </summary>
        /// <param name="topic"> The name of the HTML file in the CHM which contains the specified topic </param>
        public static void ShowTopic(string topic)
        {
            Help.ShowHelp(null, "help.chm", HelpNavigator.Topic, topic);
        }
    }
}
