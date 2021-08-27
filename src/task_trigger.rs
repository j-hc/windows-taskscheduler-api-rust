use crate::variant_d::to_bstr;
use bindings::Windows::Win32::Foundation::BSTR;


pub struct TaskIdleTrigger {
    pub id: BSTR,
    pub repetition_interval: BSTR,
    pub repetition_stop_at_duration_end: i16
}
impl TaskIdleTrigger {
    pub fn new(id: &str, repetition_interval: u32, repetition_stop_at_duration_end: bool) -> Self {
        let repetition_stop_at_duration_end: i16 = if repetition_stop_at_duration_end {1} else {0};
        Self {
            id: to_bstr(id),
            repetition_interval: to_bstr(&format!("PT{}M", repetition_interval)),
            repetition_stop_at_duration_end,
        }
    }
}


pub struct TaskLogonTrigger {
    pub id: BSTR,
    pub repetition_interval: BSTR,
    pub repetition_stop_at_duration_end: i16,
    pub delay: BSTR
}
impl TaskLogonTrigger{
    pub fn new(id: &str, repetition_interval: u32, repetition_stop_at_duration_end: bool, delay: u32) -> Self {
        let repetition_stop_at_duration_end: i16 = if repetition_stop_at_duration_end {1} else {0};
        Self {
            id: to_bstr(id),
            repetition_interval: to_bstr(&format!("PT{}M", repetition_interval)),
            repetition_stop_at_duration_end,
            delay: to_bstr(&format!("PT{}M", delay))
        }
    }
}
