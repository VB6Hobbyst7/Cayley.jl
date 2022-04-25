# Cayley

## Introduction
Currently (17 March 2022) The Cayley 2022 software is a work in progress, with some functionality not yet implemented. But the software is ready for initial testing in the Airbus environment.

## Installation

 * Installation should be done using the account under which the software will be used. If that account is not an administrator account, then Windows User Account Control will request the password for a separate administrator account.
 * Bear in mind that installation of Cayley 2022 will uninstall the 2017 verison of Cayley as they cannot co-exist on the same PC. So for the time being don't install onto a PC on which you wish to continue using the 2017 version of Cayley.
 * Both Julia and Microsoft Office must be installed on your PC, with Excel not running.
 * [JuliaExcel](https://github.com/PGS62/JuliaExcel.jl) must be installed on the PC.
 * [Snaketail](http://snakenest.com/snaketail/) must be installed on the PC. It's used to display progress in long-running operations, such as scenario analysis.
 * Launch Julia, then copy and paste (via right-click) the following commands into the Julia REPL:
   ```julia
   using Pkg
   Pkg.add("Revise")
   Pkg.add(url="https://github.com/SolumCayley/XVA")
   Pkg.add(url="https://github.com/SolumCayley/Cayley.jl",rev="v0.17")
   using Cayley
   Cayley.create_system_image()
   Cayley.installme()
      
   ```
   
 * The process will take about 10 minutes to run, as it includes a step (`Cayley.make_system_image()`) that "pre-compiles" the Julia code. At the end of the process the command `Cayley.installme()` installs the Excel components (the workbooks and Excel addins) and at that stage you will need to click 'OK' in the two dialogs shown below.
 
![image](https://user-images.githubusercontent.com/18028484/158453670-b64fcd0f-56aa-4d6d-aaf2-73de7f5d8983.png)  
 ![image](https://user-images.githubusercontent.com/18028484/158453787-1f6d92e7-3068-4080-aa37-92cf8153702d.png)

 ## First Use
 
 Here are some steps as an initial test that the software is working at Airbus.
 
  * Launch Excel
  * Open workbook `C:\ProgramData\Solum\Workbooks\Cayley2022.xlsm`
  * Review the contents of the Config worksheet. On first installation these settings should be such that the software will "work" but the settings will probably need to be amended in order that data flows (trade data, lines data and market data) work correctly in the Airbus environment. More details on all this to follow.
  * Switch to the `CreditUsage` worksheet.
  * Click the `Menu` button (or hit F8)
  * Select the `Open Trades, Market and Lines workbooks` option. Should take about 10 seconds to run (see messages in the Excel status bar at the bottom of the screen)
  * 'Menu' -> `Calculate PFE`. After a few seconds, the worksheet should look something like this:
  
  ![image](https://user-images.githubusercontent.com/18028484/158784423-60d85b74-a0e3-44ad-b0e6-9baf860cf45c.png)

 But alternatively there might be pop-up boxes containing error messages.  
 Please provide feedback to Philip Swannell  
 philip.swannell@solum-financial.com
