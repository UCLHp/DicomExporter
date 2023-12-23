# DicomExporter
ESAPI scripts for batch dicom export


### Prerequisites

Before running any scripts you will need to ensure that:  
  * You have created a dicom daemon on the Varian server and set up trusted entities as described in Chapter 4 [here](https://varianapis.github.io/VarianApiBook.pdf)
  * Created a receiver dicom daemon where you will export the dicom files to - you may wish to use a free service such as [ClearCanvas](http://clearcanvas.github.io/)
  * You know the relevant AETitles, IP addresses and port numbers for the above
  
### About the code
---
ESAPI does not provide any built-in functions for dealing with dicom and so we must use an external library. Here we've used [EvilDicom](http://rexcardan.github.io/Evil-DICOM/) by Rex Cardan which can be simply installed in Visual Studios via NuGet as described in Chapter 3 [here](https://varianapis.github.io/VarianApiBook.pdf).  

Since this code will open multiple patients it must be built as a stand-alone executable.

The script as it stands requires a 2 column csv file, where the first column is a list of Patient IDs and the second is a list of Course IDs. This file must be located in the same directory as the executable. The script will then export every plan inside the specified course if it has a valid dose. Along with each plan, the dicom files for the structure set, plan dose, plan image and all 3D images that share the same FrameOfReferenceUID will be exported (frame of reference is used so that all phases and reconstructed images such as MIP, Ave, etc. belonging to the related 4DCTs will be exported).

#### Running the code
1. Use the Eclipse Script Wizard to generate a Stand-alone executable  
2. Copy the code in the .cs file into your own
3. Install EvilDicom via NuGet and ensure the Varian libraries are referenced  
4. Update the parameters relating to the csv input file and all AETitles, IPs and ports
5. Build the solution to generate your executable
6. Ensure the executable and patient file are on Varian's server (\\vaimg103\...)
7. From with the External Beam workspace select Tools > Scripts and in the Location section select System Scripts > Open Folder. Navigate to and select your executable




