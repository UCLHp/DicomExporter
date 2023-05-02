using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using EvilDICOM.Network;
using EvilDICOM.Core;
using EvilDICOM.Core.Helpers;
using System.IO;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {

            // ------------------------------------------------------------------------------------------------
            // Testing if from just a patient and plan ID, can I get all related dicom files?
            // Get all relevant UIDs

            string patientID = "zzzPhysicsSpinePhantom";
            string courseID = "2D3D_Roll_Issue";
            string planID = "TestPlan";


            Patient patient = context.Patient;

            Course course = patient.Courses.First(x => x.Id == courseID);
            IonPlanSetup plan = course.IonPlanSetups.First(x => x.Id == planID);

            // Get all UIDs
            string planUID = plan.UID;                                         // SOP Instance UID
            string structureSetUID = plan.StructureSet.UID;                    // SOP Instance UID
            //string imageUID_zz = plan.StructureSet.Image.UID;                // This was empty
            string imageSeriesUID = plan.StructureSet.Image.Series.UID;        // Series instance UID within CT dicom files
            string imageStudyUID = plan.StructureSet.Image.Series.Study.UID;   // Study Instance UID within CT dicom files
            string doseUID = plan.Dose.UID;

            MessageBox.Show("Plan UID = " + planUID + "\nStructureSet UID = " + structureSetUID +
                "\nImageSeriesUID = " + imageSeriesUID + "\nImageStudyUID = " + imageStudyUID + "\nDoseUID = " + doseUID);


            // --------------------------------------------------------------------------------------------------
            // Only export files matching the above UIDs

            // Store the details of the daemon ( Ae Title , IP , port )
            var daemon = new Entity("ESAPI_Export", "9.140.36.85", 51801);

            // Store the details of the client ( Ae Title , port ) -> IP address is determined by CreateLocal() method
            var local = Entity.CreateLocal("VMSFSD", 104);

            var ip = local.IpAddress.ToString();
            MessageBox.Show($"IP of local = {ip}");

            // Set up a client ( DICOM SCU = Service Class User )
            var client = new DICOMSCU(local);

            // TRY C - ECHO
            var canPing = client.Ping(daemon);
            MessageBox.Show($"Ping attempt = {canPing.ToString()}");

            // Build a finder class to help with C - FIND operations
            var finder = client.GetCFinder(daemon);
            var studies = finder.FindStudies(patientID);  // Search by Patient ID
            var series = finder.FindSeries(studies);
            var images = finder.FindImages(series);

            // Write results to console
            string msg2 = $" DICOM C-Find from { local.AeTitle } => " + $" { daemon.AeTitle } @ { daemon.IpAddress }:{ daemon.Port }: ";
            msg2 += $"\n\t { studies.Count() } Studies Found ";
            msg2 += $"\n\t { series.Count() } Series Found ";
            msg2 += $"\n\t { images.Count() } Images Found ";
            MessageBox.Show(msg2);

            // Note that any dicom file is by dicom convention, an "image"; i.e. plans, dose files, structure sets, etc.
            // Filter series by modality , then create list of
            var plans = series.Where(s => s.Modality == "RTPLAN").SelectMany(ser => finder.FindImages(ser));
            var doses = series.Where(s => s.Modality == "RTDOSE").SelectMany(ser => finder.FindImages(ser));
            var cts = series.Where(s => s.Modality == "CT").SelectMany(ser => finder.FindImages(ser));
            var structsets = series.Where(s => s.Modality == "RTSTRUCT").SelectMany(ser => finder.FindImages(ser));

            string msg3 = $"Plans found = {plans.Count()}";
            msg3 += $"\nDoses found = { doses.Count() }";
            msg3 += $"\nCTs found =  { cts.Count() }";
            msg3 += $"\nStructure sets =  { structsets.Count() }";
            MessageBox.Show(msg3);


            var mover = client.GetCMover(daemon);
            ushort msgId = 1;



            string msg = "";
            foreach (var p in plans)
            {
                if (p.SOPInstanceUID == planUID)
                {
                    msg += $"\nSending plan { p.SOPInstanceUID }... ";
                    var response = mover.SendCMove(p, local.AeTitle, ref msgId);
                    msg += $"\nDICOM C-Move Results : ";
                    msg += $"\nNumber of Completed Operations : { response.NumberOfCompletedOps }";
                    msg += $"\nNumber of Failed Operations : { response.NumberOfFailedOps } ";
                    msg += $"\nNumber of Remaining Operations : { response.NumberOfRemainingOps } ";
                    msg += $"\nNumber of Warning Operations : { response.NumberOfWarningOps}";
                }

            }
            MessageBox.Show(msg);

            msg = "";
            foreach (var dose in doses)
            {
                if (dose.SOPInstanceUID == doseUID)
                {
                    msg += $"\nSending dose { dose.SOPInstanceUID }... ";
                    var response = mover.SendCMove(dose, local.AeTitle, ref msgId);
                    msg += $"\nDICOM C-Move Results : ";
                    msg += $"\nNumber of Completed Operations : { response.NumberOfCompletedOps }";
                    msg += $"\nNumber of Failed Operations : { response.NumberOfFailedOps } ";
                    msg += $"\nNumber of Remaining Operations : { response.NumberOfRemainingOps } ";
                    msg += $"\nNumber of Warning Operations : { response.NumberOfWarningOps}";
                }
            }
            MessageBox.Show(msg);

            msg = "";
            foreach (var ss in structsets)
            {
                if (ss.SOPInstanceUID == structureSetUID)
                {
                    msg += $"\nSending structure set { ss.SOPInstanceUID }... ";
                    var response = mover.SendCMove(ss, local.AeTitle, ref msgId);
                    msg += $"\nDICOM C-Move Results : ";
                    msg += $"\nNumber of Completed Operations : { response.NumberOfCompletedOps }";
                    msg += $"\nNumber of Failed Operations : { response.NumberOfFailedOps } ";
                    msg += $"\nNumber of Remaining Operations : { response.NumberOfRemainingOps } ";
                    msg += $"\nNumber of Warning Operations : { response.NumberOfWarningOps}";
                }
            }
            MessageBox.Show(msg);

            msg = "";
            foreach (var ctslice in cts)
            {
                if (ctslice.StudyInstanceUID == imageStudyUID && ctslice.SeriesInstanceUID == imageSeriesUID)
                {
                    msg += $"\nSending CT slices { ctslice.SOPInstanceUID }... ";
                    var response = mover.SendCMove(ctslice, local.AeTitle, ref msgId);
                    msg += $"\nDICOM C-Move Results : ";
                    msg += $"\nNumber of Completed Operations : { response.NumberOfCompletedOps }";
                    msg += $"\nNumber of Failed Operations : { response.NumberOfFailedOps } ";
                    msg += $"\nNumber of Remaining Operations : { response.NumberOfRemainingOps } ";
                    msg += $"\nNumber of Warning Operations : { response.NumberOfWarningOps}";
                }
            }
            MessageBox.Show(msg);





        }
    }
}
