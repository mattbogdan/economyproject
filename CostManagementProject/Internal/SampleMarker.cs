using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Research.DynamicDataDisplay.PointMarkers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;


namespace CostManagementProject.Internal
{
    class SampleMarker : ShapeElementPointMarker
    {
        public override UIElement CreateMarker()
        {
            Canvas result = new Canvas()
            {
                Width = 10,
                Height = Size
            };
            result.Width = Size;
            result.Height = Size;
            result.Background = Brush;
            if (ToolTipText != String.Empty)
            {
                ToolTip tt = new ToolTip();
                tt.Content = ToolTipText;
                result.ToolTip = tt;
            }
            return result;
        }

        public override void SetPosition(UIElement marker, Point screenPoint)
        {
            Canvas.SetLeft(marker, screenPoint.X - Size / 2);
            Canvas.SetTop(marker, screenPoint.Y - Size / 2);
        }
    }
}
