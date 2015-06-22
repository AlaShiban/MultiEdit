using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace MultiPointEdit
{
    class MultiPointEditCommandFilter : IOleCommandTarget
    {
        private IWpfTextView m_textView;
        private IAdornmentLayer m_adornmentLayer;
        private List<ITrackingPoint> m_trackList;
        private CaretPosition m_lastCaretPosition = new CaretPosition();
        private DTE2 m_dte;
        private Dictionary<string, int> positionHash = new Dictionary<string, int>();
        public MultiPointEditCommandFilter(IWpfTextView tv)
        {
            m_textView = tv;
            m_adornmentLayer = tv.GetAdornmentLayer("MultiEditLayer");
            m_trackList = new List<ITrackingPoint>();
            m_lastCaretPosition = m_textView.Caret.Position;
            m_textView.LayoutChanged += m_textView_LayoutChanged;
            m_dte = ServiceProvider.GlobalProvider.GetService(typeof(DTE)) as DTE2;
        }

        void m_textView_LayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            RedrawScreen();
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            Debug.WriteLine(nCmdID);

            bool trackListNonEmpty = m_trackList.Count > 0;
            bool performSyncedOp = false;
            if (pguidCmdGroup == typeof(VSConstants.VSStd2KCmdID).GUID)
            {
                switch (nCmdID)
                {
                    case ((uint)VSConstants.VSStd2KCmdID.TYPECHAR):
                    case ((uint)VSConstants.VSStd2KCmdID.BACKSPACE):
                    case ((uint)VSConstants.VSStd2KCmdID.DELETEWORDRIGHT):
                    case ((uint)VSConstants.VSStd2KCmdID.DELETEWORDLEFT):
                    case ((uint)VSConstants.VSStd2KCmdID.TAB):
                    case ((uint)VSConstants.VSStd2KCmdID.LEFT):
                    case ((uint)VSConstants.VSStd2KCmdID.RIGHT):
                    case ((uint)VSConstants.VSStd2KCmdID.UP):
                    case ((uint)VSConstants.VSStd2KCmdID.DOWN):
                    case ((uint)VSConstants.VSStd2KCmdID.END):
                    case ((uint)VSConstants.VSStd2KCmdID.HOME):
                    case ((uint)VSConstants.VSStd2KCmdID.PAGEDN):
                    case ((uint)VSConstants.VSStd2KCmdID.PAGEUP):
                    case ((uint)VSConstants.VSStd2KCmdID.PASTE):
                    case ((uint)VSConstants.VSStd2KCmdID.PASTEASHTML):
                    case ((uint)VSConstants.VSStd2KCmdID.BOL):
                    case ((uint)VSConstants.VSStd2KCmdID.EOL):
                    case ((uint)VSConstants.VSStd2KCmdID.RETURN):
                    case ((uint)VSConstants.VSStd2KCmdID.BACKTAB):
                    case ((uint)VSConstants.VSStd2KCmdID.WORDPREV):
                    case ((uint)VSConstants.VSStd2KCmdID.WORDNEXT):
                        performSyncedOp = trackListNonEmpty;
                        break;
                    case ((uint)VSConstants.VSStd2KCmdID.UP_EXT_COL):
                    case ((uint)VSConstants.VSStd2KCmdID.DOWN_EXT_COL):
                        // Allow expanding track list vertically but only if there isn't any prior
                        // horizontal selection. This effectively means that if the block selection is
                        // started with vertical extension we will use Multi Edit, but if it is started
                        // with horizontal extension we will use the old-school behavior.
                        performSyncedOp = m_textView.Selection.SelectedSpans.All(span => span.Length == 0);
                        break;
                    case ((uint)VSConstants.VSStd2KCmdID.LEFT_EXT):
                    case ((uint)VSConstants.VSStd2KCmdID.RIGHT_EXT):
                    case ((uint)VSConstants.VSStd2KCmdID.CANCEL):
                    case ((uint)VSConstants.VSStd2KCmdID.LEFT_EXT_COL):
                    case ((uint)VSConstants.VSStd2KCmdID.RIGHT_EXT_COL):
                        if (trackListNonEmpty)
                        {
                            // Break out of the Multi Edit mode when (unsupported) horizontal
                            // selection or ESC is used.
                            ClearSyncPoints();
                            RedrawScreen();
                        }
                        break;
                    default:
                        break;
                }
            }
            else if (pguidCmdGroup == VSConstants.GUID_VSStandardCommandSet97)
            {
                switch (nCmdID)
                {

                    case ((uint)VSConstants.VSStd97CmdID.Delete):
                    case ((uint)VSConstants.VSStd97CmdID.Paste):
                        performSyncedOp = trackListNonEmpty;
                        break;
                    default:
                        break;
                }
            }
            if (performSyncedOp)
                return SyncedOperation(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            else
                return NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        private int SyncedOperation(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            ITextCaret caret = m_textView.Caret;

            if (m_trackList.Count == 0)
                AddSyncPoint(caret.Position);

            var tempTrackList = m_trackList;
            m_trackList = new List<ITrackingPoint>();

            int result = 0;

            m_dte.UndoContext.Open("Multi-point edit");

            for (int i = 0; i < tempTrackList.Count; i++)
            {
                SnapshotPoint snapPoint = tempTrackList[i].GetPoint(m_textView.TextSnapshot);
                caret.MoveTo(snapPoint);
                Debug.Print("Caret #" + i + " pos : " + caret.Position);
                result = NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                if (pguidCmdGroup == typeof(VSConstants.VSStd2KCmdID).GUID &&
                    (nCmdID == (uint)VSConstants.VSStd2KCmdID.UP_EXT_COL || nCmdID == (uint)VSConstants.VSStd2KCmdID.DOWN_EXT_COL))
                    AddSyncPointsFromSelections();
                else
                    AddSyncPoint(m_textView.Caret.Position);
            }

            m_dte.UndoContext.Close();

            RedrawScreen();
            return result;
        }

        private void ClearSyncPoints()
        {
            m_trackList.Clear();
            m_adornmentLayer.RemoveAllAdornments();
        }

        private void RedrawScreen()
        {
            m_adornmentLayer.RemoveAllAdornments();
            positionHash.Clear();
            List<ITrackingPoint> newTrackList = new List<ITrackingPoint>();
            foreach (var trackPoint in m_trackList)
            {
                var curPosition = trackPoint.GetPosition(m_textView.TextSnapshot);
                IncrementCount(positionHash, curPosition.ToString());
                if (positionHash[curPosition.ToString()] > 1)
                    continue;
                DrawSingleSyncPoint(trackPoint);
                newTrackList.Add(trackPoint);
            }

            m_trackList = newTrackList;

        }

        private void IncrementCount(Dictionary<string, int> someDictionary, string id)
        {
            if (!someDictionary.ContainsKey(id))
                someDictionary[id] = 0;

            someDictionary[id]++;
        }

        private void DrawSingleSyncPoint(ITrackingPoint trackPoint)
        {
            if (trackPoint.GetPosition(m_textView.TextSnapshot) >= m_textView.TextSnapshot.Length)
                return;

            SnapshotSpan span = new SnapshotSpan(trackPoint.GetPoint(m_textView.TextSnapshot), 1);
            var brush = Brushes.DarkGray;
            var geom = m_textView.TextViewLines.GetLineMarkerGeometry(span);
            GeometryDrawing drawing = new GeometryDrawing(brush, null, geom);

            if (drawing.Bounds.IsEmpty)
                return;

            Rectangle rect = new Rectangle()
            {
                Fill = brush,
                Width = drawing.Bounds.Width / 6,
                Height = drawing.Bounds.Height - 4,
                Margin = new System.Windows.Thickness(0, 2, 0, 0),
            };

            Canvas.SetLeft(rect, geom.Bounds.Left);
            Canvas.SetTop(rect, geom.Bounds.Top);
            m_adornmentLayer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, "MultiEditLayer", rect, null);

        }

        private void AddSyncPoint(CaretPosition caretPosition)
        {
            CaretPosition curPosition = caretPosition;
            // We don't support Virtual Spaces [yet?]

            var curTrackPoint = m_textView.TextSnapshot.CreateTrackingPoint(curPosition.BufferPosition.Position, PointTrackingMode.Positive);
            // Check if the bounds are valid

            if (curTrackPoint.GetPosition(m_textView.TextSnapshot) >= 0)
                m_trackList.Add(curTrackPoint);
            else
            {
                curTrackPoint = m_textView.TextSnapshot.CreateTrackingPoint(0, PointTrackingMode.Positive);
                m_trackList.Add(curTrackPoint);
            }

            if (curPosition.VirtualSpaces > 0)
            {
                m_textView.Caret.MoveTo(curTrackPoint.GetPoint(m_textView.TextSnapshot));
            }
        }

        private void AddSyncPoint(int position)
        {
            m_trackList.Add(m_textView.TextSnapshot.CreateTrackingPoint(Math.Max(position, 0), PointTrackingMode.Positive));
        }

        private void AddSyncPointsFromSelections()
        {
            foreach (var span in m_textView.Selection.SelectedSpans)
                AddSyncPoint(span.Start.Position);
            m_textView.Selection.Clear();
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (pguidCmdGroup == typeof(VSConstants.VSStd2KCmdID).GUID)
            {
                for (int i = 0; i < cCmds; i++)
                {
                    switch (prgCmds[i].cmdID)
                    {
                        case ((uint)VSConstants.VSStd2KCmdID.TYPECHAR):
                        case ((uint)VSConstants.VSStd2KCmdID.BACKSPACE):
                        case ((uint)VSConstants.VSStd2KCmdID.DELETEWORDRIGHT):
                        case ((uint)VSConstants.VSStd2KCmdID.DELETEWORDLEFT):
                        case ((uint)VSConstants.VSStd2KCmdID.TAB):
                        case ((uint)VSConstants.VSStd2KCmdID.LEFT):
                        case ((uint)VSConstants.VSStd2KCmdID.RIGHT):
                        case ((uint)VSConstants.VSStd2KCmdID.UP):
                        case ((uint)VSConstants.VSStd2KCmdID.DOWN):
                        case ((uint)VSConstants.VSStd2KCmdID.END):
                        case ((uint)VSConstants.VSStd2KCmdID.HOME):
                        case ((uint)VSConstants.VSStd2KCmdID.PAGEDN):
                        case ((uint)VSConstants.VSStd2KCmdID.PAGEUP):
                        case ((uint)VSConstants.VSStd2KCmdID.PASTE):
                        case ((uint)VSConstants.VSStd2KCmdID.PASTEASHTML):
                        case ((uint)VSConstants.VSStd2KCmdID.BOL):
                        case ((uint)VSConstants.VSStd2KCmdID.EOL):
                        case ((uint)VSConstants.VSStd2KCmdID.RETURN):
                        case ((uint)VSConstants.VSStd2KCmdID.BACKTAB):
                        case ((uint)VSConstants.VSStd2KCmdID.WORDPREV):
                        case ((uint)VSConstants.VSStd2KCmdID.WORDNEXT):
                        case ((uint)VSConstants.VSStd2KCmdID.UP_EXT_COL):
                        case ((uint)VSConstants.VSStd2KCmdID.DOWN_EXT_COL):
                        case ((uint)VSConstants.VSStd2KCmdID.LEFT_EXT):
                        case ((uint)VSConstants.VSStd2KCmdID.RIGHT_EXT):
                        case ((uint)VSConstants.VSStd2KCmdID.CANCEL):
                        case ((uint)VSConstants.VSStd2KCmdID.LEFT_EXT_COL):
                        case ((uint)VSConstants.VSStd2KCmdID.RIGHT_EXT_COL):
                            prgCmds[i].cmdf = (uint)(OLECMDF.OLECMDF_ENABLED | OLECMDF.OLECMDF_SUPPORTED);
                            return VSConstants.S_OK;
                    }
                }
            }

            return NextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        public void HandleClick(bool addCursor)
        {
            if (addCursor && m_textView.Selection.SelectedSpans.All(span => span.Length == 0))
            {
                if (m_textView.Selection.SelectedSpans.Count == 1)
                {
                    if (m_trackList.Count == 0)
                        AddSyncPoint(m_lastCaretPosition);

                    AddSyncPoint(m_textView.Caret.Position);
                }
                else
                    AddSyncPointsFromSelections();
                RedrawScreen();
            }
            else if (m_trackList.Any())
            {
                ClearSyncPoints();
                RedrawScreen();
            }

            m_lastCaretPosition = m_textView.Caret.Position;
        }

        internal bool Added { get; set; }
        internal IOleCommandTarget NextTarget { get; set; }
    }
}
