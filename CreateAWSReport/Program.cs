using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using Amazon;
using Amazon.CloudWatch;
using Amazon.CloudWatch.Model;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;
using Amazon.EC2;
using Amazon.EC2.Model;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Font = iTextSharp.text.Font;

namespace CreateAWSReport
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            CreateAwsReport();
        }

        public static void CreateAwsReport()
        {
            //Create the Amazon objects that are necessary to query EC2 Instance information, and the metric data associated with these instances
            AmazonCloudWatch amazonCloudWatchClient =
                AWSClientFactory.CreateAmazonCloudWatchClient(RegionEndpoint.USEast1); //Specify the Region that is hosting your Instances
            AmazonEC2 amazonEc2 = AWSClientFactory.CreateAmazonEC2Client(RegionEndpoint.USEast1);
            
            //Create an request for the DescribeInstances Method
            DescribeInstancesRequest descInsReq = new DescribeInstancesRequest();
            
            //Create a filter that will only return running instances
            Filter runningFilter = new Filter {Name = "instance-state-name"};
            List<string> filterValues = new List<string> {"running"};
            runningFilter.Value = filterValues;
            List<Filter> filterList = new List<Filter> {runningFilter};
            descInsReq.Filter = filterList;

            DescribeInstancesResponse describeInstancesResponse = amazonEc2.DescribeInstances(descInsReq);


            //Create the objects that are used to create and populate the PDF
            Document pdfDocument = null;
            MemoryStream pdfStream = null;
            PdfWriter writer = null;

            //Loop through all the instances returned by the DescribeInstance Method
            foreach (Reservation reservation in describeInstancesResponse.DescribeInstancesResult.Reservation)
            {
                //The Name Tag for the EC2 instance
                string friendlyName = reservation.RunningInstance[0].Tag[0].Value;
                //InstanceID
                string instanceId = reservation.RunningInstance[0].InstanceId;

                //The instance dimesion tells the request which EC2 instance metrics need to be returned
                Dimension instanceDimesion = new Dimension {Name = "InstanceId", Value = instanceId};
                List<Dimension> metricDimesions = new List<Dimension> {instanceDimesion};

                GetMetricStatisticsRequest getCpuUtilizationMetricRequest = new GetMetricStatisticsRequest
                    {
                        StartTime = DateTime.UtcNow.Date.AddDays(-1), 
                        EndTime = DateTime.UtcNow.Date,
                        MetricName = "CPUUtilization",                  
                        Namespace = "AWS/EC2",                          
                        Unit = "Percent",
                        Period = 60*60,
                        Statistics = new List<string> {"Average"},
                        Dimensions = metricDimesions
                    };
                GetMetricStatisticsRequest getPhyMemUsageMetricReq = new GetMetricStatisticsRequest
                    {
                        StartTime = DateTime.UtcNow.Date.AddDays(-1),
                        EndTime = DateTime.UtcNow.Date,
                        MetricName = "PhysicalMemoryUtilization",
                        Namespace = "System/Windows",
                        Unit = "Percent",
                        Period = 60*60,
                        Statistics = new List<string> {"Average"},
                        Dimensions = metricDimesions
                    };


                GetMetricStatisticsResponse getMetricStatisticsResponse =
                    amazonCloudWatchClient.GetMetricStatistics(getCpuUtilizationMetricRequest);
                GetMetricStatisticsResult getMetricStatisticsResult =
                    getMetricStatisticsResponse.GetMetricStatisticsResult;
                
                //The Datapoints returned by the by the request are not ordered by date so they have to be sorted
                IList<Datapoint> datapoints =
                    getMetricStatisticsResult.Datapoints.OrderBy(x => x.Timestamp).ToList();

                GetMetricStatisticsResponse getPhysMemResponse =
                    amazonCloudWatchClient.GetMetricStatistics(getPhyMemUsageMetricReq);
                GetMetricStatisticsResult getPhysMemResult = getPhysMemResponse.GetMetricStatisticsResult;
                IList<Datapoint> phyMemDatapoints = getPhysMemResult.Datapoints.OrderBy(x => x.Timestamp).ToList();

                Stream cpuImgStream = null, phyMemImgStream = null;
                if (datapoints.Count > 0)
                {
                    cpuImgStream = GenerateChart("CPU Utilization", datapoints);
                    cpuImgStream.Seek(0, SeekOrigin.Begin); //Reset the stream to the beginning
                }
                if (phyMemDatapoints.Count > 0)
                {
                    phyMemImgStream = GenerateChart("Physical Memory Utilization", phyMemDatapoints);
                    phyMemImgStream.Seek(0, SeekOrigin.Begin);
                }
                List<Stream> imgStreams = new List<Stream>();
                if(cpuImgStream != null)
                    imgStreams.Add(cpuImgStream);
                if(phyMemImgStream != null)
                    imgStreams.Add(phyMemImgStream);
                //Create the PDF Document if it doesn't exist
                if (pdfDocument == null)
                    CreatePdfReport("Instance Resource Report", out writer, out pdfDocument, out pdfStream);

                AddToPdfReport(pdfDocument,friendlyName,imgStreams);
                
            }
            //After the loop close the writer and pdf stream and send the email
            writer.CloseStream = false;
            pdfDocument.Close();
            pdfStream.Seek(0, SeekOrigin.Begin);
            

            //Now send the email using the gmail account with the PDF Attached
            var fromAddress = new MailAddress("<From Address Here>");
            var toAddress = new MailAddress("<To Address Here>");

            const string fromPassword = "<From Account Password Here>";
            string subject = "AWS Resource Report " + DateTime.Today.ToString("MM/dd/yyyy");
            const string body = "PDF Instance Resource Report is Attached";

            var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                };

            using (var message = new MailMessage
                {
                    Subject = subject,
                    Body = body,
                    From = fromAddress
                })
            {
                message.To.Add(toAddress);
                message.Attachments.Add(new System.Net.Mail.Attachment(pdfStream, "AWS Status Report.pdf"));
                AlternateView plainTextView = AlternateView.CreateAlternateViewFromString(body, null, "text/plain");
                message.AlternateViews.Add(plainTextView);
                smtp.Send(message);
            }
        }
        public static void CreatePdfReport(string title, out PdfWriter writer, out Document document, out MemoryStream pdfStream)
        {

            document = new Document(PageSize.LETTER, 10, 10, 42, 35);
            pdfStream = new MemoryStream();
            writer = PdfWriter.GetInstance(document, pdfStream);          
            document.Open();
            Font titleFont = new Font(Font.FontFamily.TIMES_ROMAN, 24.0f, 0);
            Paragraph titleParagraph = new Paragraph(title, titleFont) { Alignment = 1 };
            document.Add(titleParagraph);
            document.AddTitle(title);
        }
        public static void AddToPdfReport(Document document,string headerText, List<Stream> images )
        {
            Font headerFont = new Font(Font.FontFamily.TIMES_ROMAN, 24.0f, 0);
            Paragraph header = new Paragraph(headerText, headerFont) { Alignment = 1 };
            document.Add(header);

            foreach (Stream image in images)
            {
                iTextSharp.text.Image pdfImage = iTextSharp.text.Image.GetInstance(image);
                pdfImage.ScalePercent(15f);
                //Center Align
                pdfImage.Alignment = 1;
                document.Add(pdfImage);
            }
        }
        public static Stream GenerateChart(string title,IList<Datapoint> series)
        {
            using (var ch = new Chart())
            {
                ch.AntiAliasing = AntiAliasingStyles.All;
                ch.TextAntiAliasingQuality = TextAntiAliasingQuality.High;
                //The Charts created have very low resolution and do not display very well in the PDF. To fix this I created a very large graph and
                //then scaled the graph image down before I inserted it into the generated PDF
                ch.Width = 1800;
                ch.Height = 1700;

                ChartArea chartArea = new ChartArea();
                Title chartTitle = new Title
                    {
                        Font = new System.Drawing.Font("Arial", 48),
                        Text = title
                    };
                ch.Titles.Add(chartTitle);
                
                chartArea.AxisY = new Axis
                {
                    LabelStyle = new LabelStyle
                        {
                        Font = new System.Drawing.Font("Arial", 48),
                        TruncatedLabels = false
                    },
                    
                    Minimum = 0,
                    Maximum = 100,
                    Interval = 20,
                    IsLabelAutoFit = false,
                    LabelAutoFitMaxFontSize = 48,
                    LabelAutoFitMinFontSize = 48,
                };              
                chartArea.AxisX = new Axis
                    {
                        LabelStyle = new LabelStyle
                            {
                            Font = new System.Drawing.Font("Arial", 48),
                            TruncatedLabels = false
                        },
                        Minimum = 0,
                        Maximum = 24,
                        Interval = 4,
                        IsLabelAutoFit = false,
                        LabelAutoFitMaxFontSize = 48,
                        LabelAutoFitMinFontSize = 48,
                    };
               
                ch.ChartAreas.Add(chartArea);               
                var s = new Series
                    {
                        ChartType = SeriesChartType.Line,
                        Color = Color.Blue,
                        BorderWidth = 5,
                        ShadowColor = Color.Black
                    };
                for (int j = 0; j < series.Count; j++)
                {
                    DataPoint dp = new DataPoint(j,series[j].Average);
                    s.Points.Add(dp);
                }
                ch.Series.Add(s);
                Stream outputStream = new MemoryStream();
                ch.SaveImage(outputStream,ChartImageFormat.Png);
                return outputStream;
            }
        }
    }
}