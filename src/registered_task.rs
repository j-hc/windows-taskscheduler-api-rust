use std::time::Duration;
use windows::core::Result;
use windows::Win32::Foundation::{BSTR, SYSTEMTIME};
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

fn date_to_dur(oledate: f64) -> Duration {
    let mut num: i64 = ((oledate * 86400000.0) + (if oledate >= 0.0 { 0.5 } else { -0.5 })) as i64;
    if num < 0 {
        num -= (num % 0x5265c00) * 2;
    }
    num += 0x3680b5e1fc00 - 0x3883122cd800;
    Duration::from_secs(num as u64)
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

    pub fn last_run_time(&self) -> Result<Duration> {
        let date = unsafe { self.registered_task.LastRunTime()? };
        Ok(date_to_dur(date))
    }

    pub fn last_task_result_raw(&self) -> Result<i32> {
        unsafe { self.registered_task.LastTaskResult() }
    }
    pub fn number_of_missed_runs(&self) -> Result<i32> {
        unsafe { self.registered_task.NumberOfMissedRuns() }
    }

    pub fn next_run_time(&self) -> Result<Duration> {
        let date = unsafe { self.registered_task.NextRunTime()? };
        Ok(date_to_dur(date))
    }

    pub fn xml(&self) -> Result<String> {
        unsafe { self.registered_task.Xml().map(|s| s.to_string()) }
    }

    pub fn stop(&self) -> Result<()> {
        unsafe { self.registered_task.Stop(0)? };
        Ok(())
    }

    pub fn get_run_times(
        &self,
        pst_start: &SYSTEMTIME,
        pst_end: &SYSTEMTIME,
    ) -> Result<(u32, SYSTEMTIME)> {
        unsafe {
            let mut pcount = 0;
            let mut pruntimes = SYSTEMTIME::default();

            // let systime_def = SYSTEMTIME {
            //     wYear: 1601,
            //     wDay: 1,
            //     wMonth: 1,
            //     ..Default::default()
            // };
            // let systime_infinite = SYSTEMTIME {
            //     wYear: 30827,
            //     wDay: 1,
            //     wMonth: 1,
            //     ..Default::default()
            // };
            // let pst_start = pst_start.unwrap_or(&systime_def);
            // let pst_end = pst_end.unwrap_or(&systime_infinite);

            self.registered_task.GetRunTimes(
                pst_start as *const SYSTEMTIME,
                pst_end as *const SYSTEMTIME,
                &mut pcount as *mut u32,
                &mut pruntimes as *mut _ as *mut _,
            )?;
            Ok((pcount, pruntimes))
        }
    }
}
