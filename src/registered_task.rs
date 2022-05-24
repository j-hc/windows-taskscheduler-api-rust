use std::time::Duration;
use windows::core::Result;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::TaskScheduler::{
    IRegisteredTask, IRunningTask, TASK_STATE_DISABLED, TASK_STATE_QUEUED, TASK_STATE_READY,
    TASK_STATE_RUNNING, TASK_STATE_UNKNOWN,
};

pub enum TaskState {
    TaskStateUnknown,
    TaskStateDisabled,
    TaskStateQueued,
    TaskStateReady,
    TaskStateRunning,
}

pub struct RegisteredTask {
    pub(crate) registered_task: IRegisteredTask,
}
impl RegisteredTask {
    pub fn name(&self) -> Result<String> {
        unsafe { self.registered_task.Name().map(|s| s.to_string()) }
    }

    pub fn path(&self) -> Result<String> {
        unsafe { self.registered_task.Path().map(|s| s.to_string()) }
    }

    pub fn state(&self) -> Result<TaskState> {
        unsafe {
            use TaskState::*;
            self.registered_task.State().map(|s| match s {
                TASK_STATE_UNKNOWN => TaskStateUnknown,
                TASK_STATE_DISABLED => TaskStateDisabled,
                TASK_STATE_QUEUED => TaskStateQueued,
                TASK_STATE_READY => TaskStateReady,
                TASK_STATE_RUNNING => TaskStateRunning,
                _ => unreachable!(),
            })
        }
    }

    pub fn enabled(&self) -> Result<bool> {
        unsafe { self.registered_task.Enabled().map(|s| s != 0) }
    }

    pub fn set_enabled(&self, enabled: bool) -> Result<()> {
        unsafe { self.registered_task.SetEnabled(enabled as i16) }
    }

    pub fn run_raw(&self) -> Result<IRunningTask> {
        // TODO: support variants
        unsafe { self.registered_task.Run(None) }
    }

    pub fn runex_raw(&self, flags: i32, sessionid: i32, user: &str) -> Result<IRunningTask> {
        // TODO: support variants
        unsafe {
            let user = BSTR::from(user);
            self.registered_task.RunEx(None, flags, sessionid, user)
        }
    }

    pub fn last_run_time(&self) -> Result<f64> {
        // TODO: convert COM DATE (f64) to std::Duration
        unsafe { self.registered_task.LastRunTime() }
    }

    pub fn last_task_result_raw(&self) -> Result<i32> {
        unsafe { self.registered_task.LastTaskResult() }
    }
    pub fn number_of_missed_runs(&self) -> Result<i32> {
        unsafe { self.registered_task.NumberOfMissedRuns() }
    }

    pub fn next_run_time(&self) -> Result<Duration> {
        unsafe {
            let nrt = self.registered_task.NextRunTime()? as u64;
            Ok(Duration::from_millis(nrt))
        }
    }
    pub fn xml(&self) -> Result<String> {
        unsafe { self.registered_task.Xml().map(|s| s.to_string()) }
    }

    // pub fn definiton() -> () {}
    // pub fn get_instances() -> () {}
    // pub fn get_security_descriptor() -> () {}
    // pub fn set_security_descriptor() -> () {}
    // pub fn stop() -> () {}
    // pub fn get_run_times() -> () {}
}
