
# Vehicle Information Lookup Tool


This tool helps look up information about vehicles from the National Highway Traffic Safety Administration's online database with one or more provided VIN numbers, and is primarily intended for batch processing.

This software is free (including commercial use) and open source and is provided under the terms of the [MIT License](https://opensource.org/licenses/MIT).

Features:
 * Designed for batch processing
 * Uses multiple concurrent connections
 * Tolerates network interruptions
 * Reads VIN numbers from Excel, CSV, or paste from clipboard
 * Saves results to Excel, CSV, or copy to clipboard
 * Re-orders result rows to match the order of the source data
 * Allows choosing which columns to include in the output
 * Optionally auto-corrects VIN numbers
 * Optionally discards invalid VIN numbers


  <br/>

### How to Use

To get started:
 1. Download the [current release](https://github.com/kentpmckinney/kpm-vin-lookup-tool/releases) from GitHub
 1. Run the installer
 1. After the install is complete, launch the application from the Start menu or the icon on the Desktop
 1. Most pages in the application have a Help button which provides further information for that page


<br/>

### Previewing this Project

![Screenshot](http://kentpmckinney.github.io/kpm-vin-lookup-tool/Resources/vinlookup.gif)

<br/>

### Technologies Used

  <code>C#
WPF</code>
  <br/>
  <br/>

### Working with the Source Code

<details>
  <summary>Instructions</summary>

  <br>
  The following are suggestions to help set up a development environment for this project. The actual steps needed may differ slightly depending on the operating system and other factors.

  <br/>
  <br/>

  ### Prerequisites

  The following software must be installed and properly configured on the target machine. 

   * Git (recommended)
 * .NET 7.2 or Higher
 * Visual Studio 2019
 * Windows Operating System
  <br/>

  ### Setting up a Development Environment

  The following steps are meant to be a quick way to get the project up and running.

  1. Download a copy of the source code from: https://github.com/kentpmckinney/kpm-vin-lookup-tool or clone using the repository link: https://github.com/kentpmckinney/kpm-vin-lookup-tool.git
1. Open Visual Studio 2019
1. Navigate to the folder location of the source files
1. Open the solution file
1. Press F5 to build and run
  <br/>

  ### Notes

  To gain the ability to move items around in the XAML GUI interface, look for this line:
	<code><Setter Property="Visibility" Value="Collapsed"/></code>
and set <code>Value="Visible"</code>

  ### Deployment

  In Visual Studio, under Project > Properties, set the build configuration to Release and perform a build. Program files will appear in the release folder and can be used as-is or bundled in an installation package.

</details>

<br/>

### Authors

[kentpmckinney](https://github.com/kentpmckinney)
<br/>
<br/>

### Acknowledgments

<sub>[Excel Data Reader](https://github.com/ExcelDataReader/ExcelDataReader), [EPPlus](https://github.com/JanKallman/EPPlus)</sub>
<br/>
<br/>

###### <sub>Copyright&copy; 2020 [kentpmckinney](https://github.com/kentpmckinney). All rights reserved.</sub>