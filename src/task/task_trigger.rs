use std::time::Duration;
use windows::Win32::Foundation::BSTR;

pub struct TaskIdleTrigger {
    pub(crate) id: BSTR,
    pub(crate) repetition_interval: BSTR,
    pub(crate) repetition_stop_at_duration_end: i16,
    pub(crate) execution_time_limit: BSTR,
}
impl TaskIdleTrigger {
    pub fn new(
        id: &str,
        repetition_interval: Duration,
        repetition_stop_at_duration_end: bool,
        execution_time_limit: Duration,
    ) -> Self {
        Self {
            id: id.into(),
            repetition_interval: format!("PT{}S", repetition_interval.as_secs()).into(),
            repetition_stop_at_duration_end: repetition_stop_at_duration_end as i16,
            execution_time_limit: format!("PT{}S", execution_time_limit.as_secs()).into(),
        }
    }
}

pub struct TaskLogonTrigger {
    pub(crate) id: BSTR,
    pub(crate) repetition_interval: BSTR,
    pub(crate) repetition_stop_at_duration_end: i16,
    pub(crate) execution_time_limit: BSTR,
    pub(crate) delay: BSTR,
}
impl TaskLogonTrigger {
    pub fn new(
        id: &str,
        repetition_interval: Duration,
        repetition_stop_at_duration_end: bool,
        execution_time_limit: Duration,
        delay: Duration,
    ) -> Self {
        Self {
            id: id.into(),
            repetition_interval: format!("PT{}S", repetition_interval.as_secs()).into(),
            repetition_stop_at_duration_end: repetition_stop_at_duration_end as i16,
            execution_time_limit: format!("PT{}S", execution_time_limit.as_secs()).into(),
            delay: format!("PT{}S", delay.as_secs()).into(),
        }
    }
}
