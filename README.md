The purpose of this plugin is to estimate the time a pattern feature takes to fully execute in Siemens NX. The estimation is made dynamically via a GUI Block. The plug-in is provided as a tool for designers to use by the PLM Center at Clemson University under the terms of the attached license.


https://github.com/Clemson-PLMC/PatternEstimatorTool/assets/130243567/2fd8d197-80e9-4737-972a-a6b2f13d5c55


## To Use
There are two parts to the plugin: the XML for the GUI ([PatterTimeEstimator.dlx](PatterTimeEstimator.dlx)) and the actual processing code, found in [PatternTimeEstimatorTool.vb](PatternTimeEstimatorTool.vb). While you can place the second file wherever you can access it (and where it can locate the GUI file), you'll need to place the GUI file in a directory accessible to NX. Information for how to do so is included in the NX Open Programmer's Guide (Executing NX Open Automation/[Application Directory Structure](https://docs.sw.siemens.com/en-US/doc/289054037/PL20190702084816205.nxopen_prog_guide/genid_application_root_directory_48_1916)), available via the [Siemens support center](https://support.sw.siemens.com/en-US/). Non proprietary information can be found in the [Getting Start with NX Open guide](https://docs.plm.automation.siemens.com/data_services/resources/nx/1872/nx_api/common/en_US/graphics/fileLibrary/nx/nxopen/NXOpen_Getting_Started.pdf).

This plugin was written using NXOpen, the Application Programmer Interface (API) for Siemens NX, in VB.NET. It was written for NX 1980 but can purportedly be employed in any version of NX after NX 5.0.2.
