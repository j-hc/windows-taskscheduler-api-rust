macro_rules! generate_triggers {
    ($t:ty, $name:ident, $tc:ident) => {
        struct $name {
            inner: $t,
        }
        impl $name {
            fn new(triggers: ::windows::Win32::System::TaskScheduler::ITriggerCollection) -> Self {
                let inner = unsafe {
                    let trigger = triggers
                        .Create(::windows::Win32::System::TaskScheduler::$tc)
                        .unwrap();
                    trigger.cast::<$t>().unwrap()
                };
                Self { inner }
            }

            pub fn r#type(&self) -> ::windows::Win32::System::TaskScheduler::TASK_TRIGGER_TYPE2 {
                let mut ttt =
                    ::windows::Win32::System::TaskScheduler::TASK_TRIGGER_TYPE2::default();
                unsafe {
                    self.inner
                        .Type(
                            &mut ttt
                                as *mut ::windows::Win32::System::TaskScheduler::TASK_TRIGGER_TYPE2,
                        )
                        .unwrap()
                };
                ttt
            }
        }
    };
}
generate_triggers!(
    windows::Win32::System::TaskScheduler::IIdleTrigger,
    IdleTrigger,
    TASK_TRIGGER_IDLE
);

generate_triggers!(
    windows::Win32::System::TaskScheduler::ILogonTrigger,
    LogonTrigger,
    TASK_TRIGGER_LOGON
);
generate_triggers!(
    windows::Win32::System::TaskScheduler::ITimeTrigger,
    TimeTrigger,
    TASK_TRIGGER_TIME
);
