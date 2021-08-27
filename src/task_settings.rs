use crate::variant_d::to_bstr;
use bindings::Windows::Win32::Foundation::BSTR;


pub struct IdleSettings {
    pub stop_on_idle_end: i16,
    pub restart_on_idle: i16,
    pub idle_duration: BSTR,
    pub wait_timeout: BSTR
}
impl IdleSettings {
    pub fn new(stop_on_idle_end: bool, restart_on_idle: bool, idle_duration: &str, wait_timeout: &str) -> Self {
        let stop_on_idle_end: i16 = if stop_on_idle_end {1} else {0};
        let restart_on_idle: i16 = if restart_on_idle {1} else {0};
        Self {
            stop_on_idle_end,
            restart_on_idle,
            idle_duration: to_bstr(idle_duration),
            wait_timeout: to_bstr(wait_timeout)
        }
    }
}

pub struct TaskSettings {
    pub idle_settings: IdleSettings,
    pub run_only_if_idle: i16,
    pub wake_to_run: i16,
    pub execution_time_limit: BSTR,
    pub disallow_start_if_on_batteries: i16
}
impl TaskSettings {
    pub fn new(idle_settings: IdleSettings,
           run_only_if_idle: bool,
           wake_to_run: bool,
           execution_time_limit: &str,
           disallow_start_if_on_batteries: bool)
        -> Self {
        let run_only_if_idle: i16 = if run_only_if_idle {1} else {0};
        let wake_to_run: i16 = if wake_to_run {1} else {0};
        let disallow_start_if_on_batteries: i16 = if disallow_start_if_on_batteries {1} else {0};

        Self {
            idle_settings,
            run_only_if_idle,
            wake_to_run,
            execution_time_limit: to_bstr(execution_time_limit),
            disallow_start_if_on_batteries
        }
    }
}