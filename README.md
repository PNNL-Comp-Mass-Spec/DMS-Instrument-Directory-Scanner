# Instrument Directory Scanner

This program examines the source folder on instrument computers to find files and folders.
It creates a tab-delimited text file for each instrument, listing the names of the files and folders found.
The text files for the instruments are stored in a central location, and are made
made visible via the DMS website to allow instrument operators to see the
instrument data files on the instrument computer.

## Example Instrument Source File Content

Header line:\
Folder: \\VOrbiETD03.bionet\ProteomicsData\ at 2017-11-17 10:47:18 AM

| Column   | content     | Size          | 
|----------|-------------|---------------|
| Dir      | Dispositioned | 53 GB |
| Dir      | Sequences | 24 MB |
| File     | Blank_1_5Nov17_Falcon_17-09-01.raw | 120 KB |
| File     | Blank_20170523.raw | 5.1 MB |
| File     | Blank_2_24Oct17_Falcon_17-09-02.raw | 120 KB |
| File     | QC_Shew_16_01_1_24May17_Precious_17-04-02.raw | 379 MB |
| File     | QC_Shew_16_01_2_24May17_Precious_17-04-02.raw | 408 MB |
| File     | TempSequence.sld | 2.0 KB |
| File     | TempSequence_171108140735.sld | 2.0 KB |
| File     | x_QC_Shew_17_01-500ng_2a_27Oct17_Falcon_17-09-02.raw | 760 MB |
| File     | x_QC_Shew_17_01-500ng_2a_30Oct17_Falcon_17-09-02.raw | 808 MB |

## Contacts

Written by Dave Clark and Matthew Monroe for the Department of Energy (PNNL, Richland, WA) \
E-mail: proteomics@pnnl.gov \
Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://www.pnnl.gov/integrative-omics

## License

The Instrument Directory Scanner is licensed under the 2-Clause BSD License; 
you may not use this program except in compliance with the License.
You may obtain a copy of the License at https://opensource.org/licenses/BSD-2-Clause

Copyright 2018 Battelle Memorial Institute
