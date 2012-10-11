namespace RCE.TransformManager.Tasks
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Runtime.Serialization;
    using System.Text;
    using System.Xml.Linq;

    using Microsoft.Expression.Encoder;
    using Microsoft.Web.Media.TransformManager;

    using RCE.Services.Contracts;
    using RCE.Services.Contracts.Output;

    using MediaItem = Microsoft.Expression.Encoder.MediaItem;

    public class RCEProjectTask : ITask
    {
        private bool initialized;
        private Job job;
        private ILogger transformLogger;
        private IJobMetadata transformMetadata;
        private ITaskStatus transformTaskStatus;

        private bool disposed;

        public void Initialize(ITaskStatus status, IJobMetadata metadata, ILogger logger)
        {
            if (status == null)
            {
                throw new ArgumentNullException("status");
            }

            if (metadata == null)
            {
                throw new ArgumentNullException("metadata");
            }

            if (logger == null)
            {
                throw new ArgumentNullException("logger");
            }

            this.transformTaskStatus = status;
            this.transformMetadata = metadata;
            this.transformLogger = logger;

            this.transformLogger.WriteLine(LogLevel.Verbose, "Before job instantiation");
            this.job = new Job();
            this.transformLogger.WriteLine(LogLevel.Verbose, "After job instantiation");
            this.AssignJobParameters();
            this.initialized = true;
        }

        public void Start()
        {
            if (!this.initialized)
            {
                throw new TypeInitializationException("RCE.TransformManager.Tasks.RCEProjectTask", null);
            }

            this.transformLogger.WriteLine(LogLevel.Information, "Begin encode.");

            try
            {
                this.job.Encode();
                this.transformTaskStatus.UpdateStatus(100, JobStatus.Finished, null);
            }
            catch (Exception innerException)
            {
                this.transformLogger.WriteLine(LogLevel.Error, "Caught an Exception while encoding media");
                while (innerException != null)
                {
                    this.transformLogger.WriteLine(LogLevel.Error, innerException.GetType().FullName);
                    if (!string.IsNullOrEmpty(innerException.Message))
                    {
                        this.transformLogger.WriteLine(LogLevel.Error, innerException.Message);
                    }

                    if (!string.IsNullOrEmpty(innerException.StackTrace))
                    {
                        this.transformLogger.WriteLine(LogLevel.Error, innerException.StackTrace);
                    }

                    innerException = innerException.InnerException;
                }

                this.transformTaskStatus.UpdateStatus(100, JobStatus.Failed, null);
            }

            this.transformLogger.WriteLine(LogLevel.Information, "End encode.");
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if ((!this.disposed && disposing) && (this.job != null))
            {
                this.job.Dispose();
            }

            this.disposed = true;
        }

        private void AssignJobParameters()
        {
            this.job.CreateSubfolder = false;
            this.job.SaveJobFileToOutputDirectory = false;
            this.job.EncodeProgress += this.OnProgress;
            this.job.OutputDirectory = this.transformMetadata.OutputFolder;
            this.transformLogger.WriteLine(LogLevel.Verbose, "Output directory: " + this.transformMetadata.OutputFolder);
            
            string preset = null;
            IManifestProperty property = this.transformMetadata.GetProperty(RCENamespaces.RCE + "preset");
            
            if (property != null)
            {
                preset = property.Value;
                if (!string.IsNullOrEmpty(preset))
                {
                    preset = Environment.ExpandEnvironmentVariables(preset);
                    this.transformLogger.WriteLine(LogLevel.Information, "Preset file: " + preset);
                }
            }
            
            IEnumerable<XElement> elements = this.transformMetadata.Manifest.Descendants("ref").Select(el => el);
            
            foreach (XElement element in elements)
            {
                Exception exception;
                try
                {
                    string fileName = (string)element.Attribute("src");
                    this.transformLogger.WriteLine(LogLevel.Information, "Transforming file: " + fileName);

                    this.transformLogger.WriteLine(
                        LogLevel.Information, "Deserializing file into RCE project: " + fileName);

                    string projectXml;
                    using (StreamReader reader = new StreamReader(fileName))
                    {
                        projectXml = reader.ReadToEnd();
                    }

                    Project project = this.Deserialize<Project>(projectXml);

                    this.transformLogger.WriteLine(LogLevel.Information, "Looking for RCE Visual Track");

                    Sequence sequence = project.Sequences.SingleOrDefault();

                    if (sequence != null)
                    {
                        Track track = sequence.Tracks.SingleOrDefault(t => t.TrackType.ToUpperInvariant() == "VISUAL");

                        if (track.Shots.Count > 0)
                        {
                            this.transformLogger.WriteLine(LogLevel.Information, "Adding RCE Shots");

                            Shot shot = track.Shots[0];

                            this.transformLogger.WriteLine(LogLevel.Information, "Adding RCE Shot: " + shot.Title);

                            string videoPath = this.ExtractVideoPath(shot);

                            MediaItem item = new MediaItem(videoPath)
                                { OutputFileName = Guid.NewGuid() + ".{Default Extension}" };

                            item.Sources[0].Clips[0].StartTime =
                                TimeSpan.FromSeconds(shot.SourceAnchor.MarkIn.GetValueOrDefault());

                            this.transformLogger.WriteLine(
                                LogLevel.Information, "Start Time: " + item.Sources[0].Clips[0].StartTime);

                            item.Sources[0].Clips[0].EndTime =
                                TimeSpan.FromSeconds(shot.SourceAnchor.MarkOut.GetValueOrDefault());

                            this.transformLogger.WriteLine(
                                LogLevel.Information, "End Time: " + item.Sources[0].Clips[0].EndTime);

                            if (track.Shots.Count > 1)
                            {
                                for (int i = 1; i < track.Shots.Count; i++)
                                {
                                    shot = track.Shots[i];
                                    this.transformLogger.WriteLine(
                                        LogLevel.Information, "Adding RCE Shot: " + shot.Title);

                                    videoPath = this.ExtractVideoPath(shot);

                                    Source source = new Source(videoPath);
                                    source.Clips[0].StartTime =
                                        TimeSpan.FromSeconds(shot.SourceAnchor.MarkIn.GetValueOrDefault());

                                    this.transformLogger.WriteLine(
                                        LogLevel.Information, "Start Time: " + source.Clips[0].StartTime);

                                    source.Clips[0].EndTime =
                                        TimeSpan.FromSeconds(shot.SourceAnchor.MarkOut.GetValueOrDefault());
                                    item.Sources.Add(source);

                                    this.transformLogger.WriteLine(
                                        LogLevel.Information, "End Time: " + source.Clips[0].EndTime);
                                }
                            }

                            OutputMetadata metadata = project.Metadata as OutputMetadata;

                            if (metadata != null)
                            {
                                this.transformLogger.WriteLine(LogLevel.Information, "Applying RCE Metadata");
                                if (!string.IsNullOrEmpty(metadata.Settings.ResizeMode))
                                {
                                    switch (metadata.Settings.ResizeMode)
                                    {
                                        case "Stretch":
                                            item.VideoResizeMode = VideoResizeMode.Stretch;
                                            break;
                                        case "Letterbox":
                                            item.VideoResizeMode = VideoResizeMode.Letterbox;
                                            break;
                                    }
                                }
                            }

                            if (!string.IsNullOrEmpty(preset))
                            {
                                this.transformLogger.WriteLine(LogLevel.Information, "Applying preset");
                                item.ApplyPreset(preset);
                            }

                            this.transformLogger.WriteLine(LogLevel.Information, "Adding media item to job");

                            try
                            {
                                this.job.MediaItems.Add(item);
                            }
                            catch (Exception exception1)
                            {
                                exception = exception1;
                                this.transformLogger.WriteLine(LogLevel.Information, exception.Message);
                                if (exception.InnerException != null)
                                {
                                    this.transformLogger.WriteLine(
                                        LogLevel.Information, exception.InnerException.Message);
                                }
                            }
                        }
                    }
                }
                catch (Exception exception2)
                {
                    exception = exception2;
                    this.transformLogger.WriteLine(LogLevel.Error, "Exception");
                    this.transformLogger.WriteLine(LogLevel.Error, exception.StackTrace);
                    this.transformLogger.WriteLine(LogLevel.Error, exception.Message);
                }
            }
        }

        private string ExtractVideoPath(Shot shot)
        {
            string shotRef = shot.Source.Resources[0].Ref;

            if (shotRef.EndsWith("/manifest"))
            {
                shotRef = shotRef.Remove(shotRef.LastIndexOf("/", StringComparison.OrdinalIgnoreCase));
            }

            string videoFileName = shotRef.Substring(shotRef.LastIndexOf("/", StringComparison.OrdinalIgnoreCase) + 1);

            if (videoFileName.EndsWith(".ism"))
            {
                string ismvPattern = string.Concat(Path.GetFileNameWithoutExtension(videoFileName), "*.ismv");

                this.transformLogger.WriteLine(LogLevel.Information, string.Format("Looking for ismv files with pattern {0} on {1}", ismvPattern, this.transformMetadata.InputFolder));

                string[] files = Directory.GetFiles(this.transformMetadata.InputFolder, ismvPattern);

                if (files.Length > 0)
                {
                    videoFileName = Path.GetFileName(files[0]);
                }
            }

            return Path.Combine(this.transformMetadata.InputFolder, videoFileName);
        }

        private void OnProgress(object sender, EncodeProgressEventArgs e)
        {
            ushort percentComplete = (ushort)(((100 / e.TotalPasses) * (e.CurrentPass - 1)) + (((ushort)e.Progress) / e.TotalPasses));
            this.transformLogger.WriteLine(LogLevel.Information, string.Concat(new object[] { "Pass ", e.CurrentPass, " of ", e.TotalPasses, " on ", e.CurrentItem.Name }));
            this.transformLogger.WriteLine(LogLevel.Information, string.Concat(new object[] { "Pass progress = ", e.Progress.ToString(CultureInfo.InvariantCulture), "; total progress = ", percentComplete }));
            this.transformTaskStatus.UpdateStatus(percentComplete, JobStatus.Running, null);
        }

        /// <summary>
        /// Deserializes the result into a known type.
        /// </summary>
        /// <typeparam name="T">The known type.</typeparam>
        /// <param name="result">The result being deserialized.</param>
        /// <returns>A known type instance.</returns>
        private T Deserialize<T>(string result)
        {
            try
            {
                byte[] bytes = Encoding.UTF8.GetBytes(result);
                T graph;
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    DataContractSerializer serializer = new DataContractSerializer(typeof(T));
                    graph = (T)serializer.ReadObject(ms);
                }

                return graph;
            }
            catch (Exception exception)
            {
                this.transformLogger.WriteLine(LogLevel.Error, "Exception");
                this.transformLogger.WriteLine(LogLevel.Error, exception.StackTrace);
                this.transformLogger.WriteLine(LogLevel.Error, exception.Message);
                throw;
            }
        }
    }
}
