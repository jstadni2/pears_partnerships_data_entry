# PEARS Partnerships Data Entry

{High-level project description goes here}

## Installation

The recommended way to install the PEARS Partnerships Data Entry script is through git, which can be downloaded [here](https://git-scm.com/downloads). Once downloaded, run the following command:

```bash
git clone https://github.com/jstadni2/pears_partnerships_data_entry
```

Alternatively, this repository can be downloaded as a zip file via this link:
[https://github.com/jstadni2/pears_partnerships_data_entry/zipball/master/](https://github.com/jstadni2/pears_partnerships_data_entry/zipball/master/)

This repository is designed to run out of the box on a Windows PC using Docker and the [/example_inputs](https://github.com/jstadni2/pears_partnerships_data_entry/tree/master/example_inputs) and [/example_outputs](https://github.com/jstadni2/pears_partnerships_data_entry/tree/master/example_outputs) directories.
To run the script in its current configuration, follow [this link](https://docs.docker.com/desktop/windows/install/) to install Docker Desktop for Windows. 

With Docker Desktop installed, this script can be run simply by double-clicking the `run_script.bat` file in your local directory.

The `run_script.bat` file can also be run in Command Prompt by entering the following command with the appropriate path:

```bash
C:\path\to\pears_partnerships_data_entry\run_script.bat
```

### Setup instructions for SNAP-Ed implementing agencies

The following steps are required to execute the PEARS Partnerships Data Entry script using your organization's PEARS data:
1. Contact [PEARS support](mailto:support@pears.io) to set up an [AWS S3](https://aws.amazon.com/s3/) bucket to store automated PEARS exports.
2. Download the automated PEARS exports. Illinois Extension's method for downloading exports from the S3 is detailed in the [PEARS Nightly Export Reformatting script](https://github.com/jstadni2/pears_nightly_export_reformatting/blob/6f370389776fb8f88495fbe4e7918c203fd84997/pears_nightly_export_reformatting.py#L9-L45).
3. Set the appropriate input and output paths in `pears_partnerships_data_entry.py` and `run_script.bat`.
	- The [Input Files](#input-files) and [Output Files](#output-files) sections provide an overview of required and output data files.
	- Copying input files to the build context would enable continued use of Docker and `run_script.bat` with minimal modifications.
	- `pears_partnerships_data_entry.py` may require additional alterations depending on the staff list format. 
4. Set the username and password variables in [pears_partnerships_data_entry.py](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/pears_partnerships_data_entry.py#L764-L765) using valid Office 365 credentials.	

### Additional setup considerations

- The formatting of PEARS export workbooks changes periodically. The example PEARS exports included in the [/example_inputs](https://github.com/jstadni2/pears_partnerships_data_entry/tree/master/example_inputs) directory are based on workbooks downloaded on 08/26/22.
Modifications to `pears_partnerships_data_entry.py` may be necessary to run with subsequent PEARS exports.
- Illinois Extension utilized [Task Scheduler](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-start-page) to run this script from a Windows PC on a monthly basis.
- Plans to deploy the PEARS Partnerships Data Entry script on AWS were never implemented and are currently beyond the scope of this repository.
- Other SNAP-Ed implementing agencies intending to utilize the PEARS Partnerships Data Entry script should consider the following adjustments as they pertain to their organization:
	- {Specific project setup considerations go here}
	
## Input Files

The following input files are required to run the PEARS Partnerships Data Entry script:
- [FY22_INEP_Staff_List.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_inputs/FY22_INEP_Staff_List.xlsx): A workbook that compiles various lists of INEP staff.
- [Illinois Extension Unit Counties.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_inputs/Illinois%20Extension%20Unit%20Counties.xlsx): A workbook that maps Illinois counties to Illinois Extension units.
- Reformatted PEARS module exports output from the [PEARS Nightly Export Reformatting script](https://github.com/jstadni2/pears_nightly_export_reformatting):
    - [Coalition_Export.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_inputs/Coalition_Export.xlsx)
    - [Indirect_Activity_Export.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_inputs/Indirect_Activity_Export.xlsx)
    - [Partnership_Export.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_inputs/Partnership_Export.xlsx)
    - [Program_Activities_Export.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_inputs/Program_Activities_Export.xlsx)
    - [PSE_Site_Activity_Export.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_inputs/PSE_Site_Activity_Export.xlsx)

Example input files are provided in the [/example_inputs](https://github.com/jstadni2/pears_partnerships_data_entry/tree/master/example_inputs) directory. 
PEARS module exports included as example files are generated using the [Faker](https://faker.readthedocs.io/en/master/) Python package and do not represent actual program evaluation data. 

## Output Files

The following output files are produced by the PEARS Partnerships Data Entry script:
- [CPHP Partnerships Data Entry YYYY-MM.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_outputs/CPHP%20Partnerships%20Data%20Entry%202022-07.xlsx): A workbook that compiles the Partnerships to enter based on CPHP PEARS records.
- [SNAP-Ed Partnerships Data Entry YYYY-MM.xlsx](https://github.com/jstadni2/pears_partnerships_data_entry/blob/master/example_outputs/SNAP-Ed%20Partnerships%20Data%20Entry%202022-07.xlsx): A workbook that compiles the Partnerships to enter based on SNAP-Ed PEARS records


Example output files are provided in the [/example_outputs](https://github.com/jstadni2/pears_partnerships_data_entry/tree/master/example_outputs) directory.