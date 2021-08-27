fn main() {
    windows::build! {
        Windows::Win32::System::Com::{
            CoInitializeEx,
            CoInitializeSecurity,
            CoCreateInstance,
        },
        Windows::Win32::System::TaskScheduler::{
            TaskScheduler,
            ITaskService,
            ITaskFolder,
            ITaskDefinition,
            IRegistrationInfo,
            IPrincipal,
            ITaskSettings,
            IIdleSettings,
            ITriggerCollection,
            ITrigger,
            ITimeTrigger,
            IActionCollection,
            IAction,
            IRegisteredTask,
            TASK_LOGON_TYPE,
            TASK_CREATION,
            IBootTrigger,
            IRepetitionPattern,
            IExecAction,
            ILogonTrigger,
            IIdleTrigger
        },
        Windows::Win32::System::OleAutomation::{VARIANT, EXCEPINFO},
        Windows::Win32::Foundation::PWSTR,
    };
}
