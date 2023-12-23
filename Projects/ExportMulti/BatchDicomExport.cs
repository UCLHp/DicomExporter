using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;
using System.Windows.Forms;
using EvilDICOM.Network;
using EvilDICOM.Core;
using EvilDICOM.Core.Helpers;
using Application = VMS.TPS.Common.Model.API.Application;


// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]


namespace BatchExport
{

    class Program
    {
        private static bool DEBUG = false;

        // ------------- Set these parameters ---------------

        static string PATIENT_FILE = "PatientAndCourse.csv";

        // ARIA daemon details
        static string ARIA_AE_TITLE = "AETitle";
        static string ARIA_IP = "1.123.12.12";
        static int ARIA_PORT = 1234;

        //Receiver daemon details
        static string RECEIVER_AE_TITLE = "AETitle2";
        static string RECEIVER_IP = "2.234.34.345";
        static int RECEIVER_PORT = 3456;

        // Trusted local entity
        static string LOCAL_AE_TITLE = "AETitle3";
        static int LOCAL_PORT = 100;

        // ---------------------------------------------------




        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if (DEBUG)
                {
                    Execute(null);
                }
                else
                {
                    using (Application app = Application.CreateApplication())
                    {
                        Execute(app);
                    }
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }
        }


        /// <summary>
        /// This method reads in a two-column file of patient MRNs and Courses as a List 
        /// of PlanRef objects
        /// </summary>
        /// <param><c>path</c> is path to file</param>
        public static List<PlanRef> ReadPatientCourseList(string path)
        {
            List<PlanRef> patientCourseList = new List<PlanRef>();
            using (TextFieldParser parser = new TextFieldParser(path))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                while (!parser.EndOfData)
                {
                    //Processing row
                    string[] fields = parser.ReadFields();
                    PlanRef pr = new PlanRef()
                    {
                        PatientID = fields[0],
                        CourseID = fields[1],
                    };

                    patientCourseList.Add(pr);
                }
            }

            // Remove duplicate entries and return
            return patientCourseList.GroupBy(elem => new { elem.PatientID, elem.CourseID } ).Select(group => group.First()).ToList();
            //return patientCourseList.Distinct().ToList();
        }


        /// <summary>
        /// This method takes a Course and searches for all photon plans (PlanSetups) with a valid dose 
        /// Returns list of PlanSetup objects.
        /// </summary>
        /// <param><c>course</c> is an ESAPI Course object</param>
        private static List<PlanSetup> FindPlansForExport(Course course)
        {
            List<PlanSetup> plansForExport = new List<PlanSetup>();
            foreach (PlanSetup p in course.PlanSetups)
            {
                if (p.IsDoseValid)
                {
                    plansForExport.Add(p);
                }
            }
            return plansForExport;
        }



        static void Execute(Application app)
        {


            // Set up Daemon properties ( Ae Title , IP , port )
            var daemon = new Entity(ARIA_AE_TITLE, ARIA_IP, ARIA_PORT);
            var uclpc = new Entity(RECEIVER_AE_TITLE, RECEIVER_IP, RECEIVER_PORT);

            // Set up a client DICOM SCU (Service Class User); must be trusted entity
            var local = Entity.CreateLocal(LOCAL_AE_TITLE, LOCAL_PORT);
            var client = new DICOMSCU(local);

            var finder = client.GetCFinder(daemon);
            var mover = client.GetCMover(daemon);


            //Read objects to export from CSV
            //List<string> patientList = ReadPatientList(PATIENT_FILE); // Put in same directory as exe
            List<PlanRef> patientList = ReadPatientCourseList(PATIENT_FILE); // Put in same directory as exe

            foreach (var patco in patientList)
            {
                // Ensure no other patient is open
                app.ClosePatient();

                // Store all UIDs of all images exported for this patient to prevent repeated exports
                List<String> allInstanceUIDs = new List<string>();

                try
                {                  
                    Patient patient = app.OpenPatientById(patco.PatientID);
                    Course course = patient.Courses.Where(x => x.Id == patco.CourseID).First();

                    // Determine which plan(s) to export
                    List<PlanSetup> plansForExport = new List<PlanSetup>();
                    plansForExport = FindPlansForExport(course);

                    foreach (PlanSetup plan in plansForExport)
                    {

                        // Get UIDs of all related plan data to be exported
                        string planUID = plan.UID;                                         // SOP Instance UID
                        string structureSetUID = plan.StructureSet.UID;                    // SOP Instance UID
                        string imageSeriesUID = plan.StructureSet.Image.Series.UID;        // Series instance UID within CT dicom files
                        string imageStudyUID = plan.StructureSet.Image.Series.Study.UID;   // Study Instance UID within CT dicom files
                        string imageFORUID = plan.StructureSet.Image.Series.FOR;           // All 4DCT images and reconstructions should share this                                                              
                        string doseUID = plan.Dose.UID;

                        // Loop all images and store series UID of any image with same FrameOfReference as planning scan
                        //   (i.e. export all related 4DCT images)
                        List<String> seriesUidsForExport = new List<String>();
                        seriesUidsForExport.Add(imageSeriesUID);
                        foreach (var study in patient.Studies)
                        {
                            foreach(var img in study.Images3D)
                            {
                                string forUID = img.FOR;
                                if (forUID == imageFORUID)
                                {
                                    seriesUidsForExport.Add(img.Series.UID);
                                }
                            }
                        }
                        seriesUidsForExport.Distinct().ToList();


                        // Use EvilDicom to query databse
                        var studies = finder.FindStudies(patient.Id);  
                        var series = finder.FindSeries(studies);
                        //var images = finder.FindImages(series);

                        // Any dicom file is by dicom convention an "image"; i.e. plans, dose files, structure sets, etc.
                        // Filter series by modality , then create list of
                        var plans = series.Where(s => s.Modality == "RTPLAN").SelectMany(ser => finder.FindImages(ser));
                        var doses = series.Where(s => s.Modality == "RTDOSE").SelectMany(ser => finder.FindImages(ser));
                        var cts = series.Where(s => s.Modality == "CT").SelectMany(ser => finder.FindImages(ser));
                        var structsets = series.Where(s => s.Modality == "RTSTRUCT").SelectMany(ser => finder.FindImages(ser));


                        try
                        {
                            // Export each dicom type sequentially; linked by UIDs to plan
                            // Store each InstanceUID that is exported
                            ushort msgId = 1;
                            foreach (var p in plans)
                            {
                                if (p.SOPInstanceUID == planUID && !allInstanceUIDs.Contains(planUID))
                                {
                                    var response = mover.SendCMove(p, uclpc.AeTitle, ref msgId);
                                    allInstanceUIDs.Add(planUID);
                                }
                            }
                            foreach (var dose in doses)
                            {
                                if (dose.SOPInstanceUID == doseUID && !allInstanceUIDs.Contains(doseUID))
                                {
                                    var response = mover.SendCMove(dose, uclpc.AeTitle, ref msgId);
                                    allInstanceUIDs.Add(doseUID);
                                }
                            }
                            foreach (var ss in structsets )
                            {
                                if (ss.SOPInstanceUID == structureSetUID && !allInstanceUIDs.Contains(structureSetUID))
                                {
                                    var response = mover.SendCMove(ss, uclpc.AeTitle, ref msgId);
                                    allInstanceUIDs.Add(structureSetUID);
                                }
                            }
                            foreach (var ctslice in cts)
                            {
                                //if (ctslice.StudyInstanceUID == imageStudyUID && ctslice.SeriesInstanceUID == imageSeriesUID && !allInstanceUIDs.Contains(ctslice.SOPInstanceUID)
                                if (seriesUidsForExport.Contains(ctslice.SeriesInstanceUID) && !allInstanceUIDs.Contains(ctslice.SOPInstanceUID))
                                {
                                    var response = mover.SendCMove(ctslice, uclpc.AeTitle, ref msgId);
                                    allInstanceUIDs.Add(ctslice.SOPInstanceUID);
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // Currently just catch all errors and crash
                            throw;
                        }

                    }
                }
                catch (Exception)
                {
                    // Currently just catch all errors and crash
                    throw;
                }


            }

            app.ClosePatient();
        }


    }


    /// <summary>
    /// Class <c>PlanRef</c> contains the string ID of a Patient and a Course
    /// </summary>
    public class PlanRef
    {
        public string PatientID = "";
        public string CourseID = "";

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
                return false;
            var xx = (PlanRef)obj;
            return (PatientID == xx.PatientID && CourseID == xx.CourseID);
        }

    }



}
