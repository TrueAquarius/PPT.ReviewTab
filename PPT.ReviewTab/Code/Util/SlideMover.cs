using System;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;


namespace PPT.ReviewTab.Code.Util
{
    public class SlideMover
    {
        private Application _powerPointApp;

        public SlideMover(Application powerPointApp)
        {
            _powerPointApp = powerPointApp;
        }

        public void MoveSelectedSlidesToEnd()
        {
            var activePresentation = _powerPointApp.ActivePresentation;
            if (activePresentation == null)
            {
                System.Windows.Forms.MessageBox.Show("No active presentation found.");
                return;
            }

            var slides = activePresentation.Slides;
            var selection = _powerPointApp.ActiveWindow.Selection;
            if (selection == null || selection.Type != PpSelectionType.ppSelectionSlides)
            {
                System.Windows.Forms.MessageBox.Show("No slides selected. Please select slides to move.");
                return;
            }

            // Get the indices of all selected slides
            var selectedSlides = selection.SlideRange.Cast<Slide>().OrderBy(slide => slide.SlideIndex).ToList();
            if (selectedSlides.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("No slides selected.");
                return;
            }

            // Find the slide index just before the first selected slide
            int previousSlideIndex = selectedSlides.First().SlideIndex - 1;

            // Move selected slides to the end
            int totalSlides = slides.Count;
            foreach (var slide in selectedSlides)
            {
                slide.MoveTo(totalSlides);
            }

            // Select and display the slide that was just before the first selected slide
            if (previousSlideIndex > 0)
            {
                slides[previousSlideIndex].Select();
                _powerPointApp.ActiveWindow.View.GotoSlide(slides[previousSlideIndex].SlideIndex);
            }
            else
            {
                // If no slide existed before the first selected slide, select and display the new first slide
                slides[1].Select();
                _powerPointApp.ActiveWindow.View.GotoSlide(slides[1].SlideIndex);
            }
        }
    }
}
