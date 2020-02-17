using System;
using System.Threading;

namespace RFIDDemoCS
{
    public enum TriggerStatus 
    { 
        NORMAL = 0, 
        UP, 
        DOWN 
    }

    public class TriggerEventArgs : EventArgs
    {
        protected TriggerStatus triggerStatus = TriggerStatus.NORMAL;

        public TriggerEventArgs(TriggerStatus triggerStatus)
        {
            this.triggerStatus = triggerStatus;
        }

        public TriggerStatus Status
        {
            get
            {
                return triggerStatus;
            }
        }
    }

    public delegate void TriggerEventHandle(object sender, TriggerEventArgs args);

    public class ScanTrigger : IDisposable
    {
        protected bool bRuning = true;
        protected ManualResetEvent triggerEvent = new ManualResetEvent(true);
        protected IntPtr hObject = IntPtr.Zero;
        public event TriggerEventHandle TriggerDown;
        //public event TriggerEventHandle TriggerUp;
        private bool disposed = false;
        
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed 
                // and unmanaged resources.
                if (disposing)
                {
                    // Dispose managed resources.
                }
                // Release unmanaged resources. If disposing is false, 
                // only the following code is executed.
                bRuning = false;
                if (hObject != IntPtr.Zero)
                {
                    CoreDLL.CloseHandle(hObject);
                }
            }
            disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public ScanTrigger()
        {
            hObject = CoreDLL.CreateEvent(IntPtr.Zero, false, false, CoreDLL.TriggerKeyEvent);
            Thread thread = new Thread(new ThreadStart(ThreadProc));
            thread.Start();
        }

        ~ScanTrigger()
        {
            Dispose(false);
        }

        public bool IsRuning
        {
            get
            {
                return bRuning;
            }
        }

        public void DoneTrigger()
        {
            triggerEvent.Set();
        }

        protected virtual void OnTriggerDown(object sender, TriggerEventArgs e)
        {
            if (TriggerDown != null)
            {
                TriggerDown(sender, e);
            }
        }

        //public virtual void OnTriggerUp(object sender, TriggerEventArgs e)
        //{
        //    if (TriggerUp != null)
        //        TriggerUp(sender, e);
        //}

        public void ThreadProc()
        {
            bool bLeftTrigger = false;
            bool bRightTrigger = false;
            TriggerStatus triggerStatus = TriggerStatus.NORMAL;

            while (bRuning)
            {
                triggerEvent.WaitOne();

                bLeftTrigger = SysIoApi.TriggerKeyStatus(SysIoApi.LEFT_TRIGGER_KEY);
                bRightTrigger = SysIoApi.TriggerKeyStatus(SysIoApi.RIGHT_TRIGGER_KEY);

                if (bLeftTrigger || bRightTrigger) //trigger key down
                {
                    triggerEvent.Reset();

                    triggerStatus = TriggerStatus.DOWN;
                    TriggerEventArgs args = new TriggerEventArgs(triggerStatus);
                    OnTriggerDown(this, args);
                }

                Thread.Sleep(100);
            }
        }
    }
}
