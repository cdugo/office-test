using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Grpc.Core;
using Microsoft.Extensions.Logging;
using PowerPointOpener;
using System.Threading.Tasks;

namespace office_test
{
    public class PowerPointService : PowerPointOpener.PowerPointBase
    {
        private readonly ILogger<PowerPointService> _logger;
        public PowerPointService(ILogger<PowerPointService> logger)
        {
            _logger = logger;
        }

        public override async Task<OpenSlideReply> OpenSlide(OpenSlideRequest request, ServerCallContext context)
        {
            // Call your OpenPowerPointAtSlide method here
            OpenPowerPointAtSlide(request.FilePath, request.SlideNumber);

            return new OpenSlideReply
            {
                Message = "PowerPoint opened successfully."
            };
        }

        // Your OpenPowerPointAtSlide method should go here
        void OpenPowerPointAtSlide(string filePath, int slideNumber)
        {
            var app = new PowerPoint.Application();
            app.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

            var pres = app.Presentations.Open(filePath, WithWindow: Microsoft.Office.Core.MsoTriState.msoTrue);

            // Navigate to the slide
            app.ActiveWindow.View.GotoSlide(slideNumber);

            // Clean up the COM objects
            Marshal.ReleaseComObject(pres);
            Marshal.ReleaseComObject(app);
        }
    }
}
