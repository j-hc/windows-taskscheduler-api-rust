use bindings::{
    Windows::Win32::System::Com::{
        CoInitializeEx,
        COINIT_MULTITHREADED,
        CoCreateInstance,
        // RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        // RPC_C_IMP_LEVEL_IMPERSONATE,
        CLSCTX_ALL
    },
    Windows::Win32::System::TaskScheduler::{
        TaskScheduler,
        ITaskService,
        // CLSID_TaskScheduler,
        ITaskFolder,
        ITaskDefinition,
        IRegistrationInfo,
        IPrincipal,
        ITaskSettings,
        IIdleSettings,
        ITriggerCollection,
        // ITrigger,
        // ITimeTrigger,
        IActionCollection,
        IAction,
        IRegisteredTask,
        // TASK_LOGON_TYPE,
        // TASK_TRIGGER_TYPE2,
        // IBootTrigger,
        // TASK_ACTION_TYPE,
        // TASK_CREATION,
        IExecAction,
        ILogonTrigger,
        IIdleTrigger,
        IRepetitionPattern,
        TASK_TRIGGER_IDLE,
        TASK_TRIGGER_LOGON,
        TASK_CREATE_OR_UPDATE,
        TASK_LOGON_INTERACTIVE_TOKEN,
        TASK_ACTION_EXEC,
        TASK_RUNLEVEL_HIGHEST,
        TASK_RUNLEVEL_TYPE
        // TASK_LOGON_SERVICE_ACCOUNT
    },
};

use windows::Interface;

use crate::variant_d::*;
use crate::task_trigger::{TaskLogonTrigger, TaskIdleTrigger};
use crate::task_action::TaskAction;
use crate::task_settings::TaskSettings;


pub enum RunLevel {
    HIGHEST
}


pub struct Task {
    name: String,
    task_service: ITaskService,
    task_definition: ITaskDefinition,
    reg_info: IRegistrationInfo,
    triggers: ITriggerCollection,
    actions: IActionCollection,
    settings: ITaskSettings,
    folder: Option<ITaskFolder>
}
impl Task {
    pub fn new(name: String) -> Result<Self, windows::Error> {
        unsafe {
            CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED)?;
            let task_service: ITaskService = CoCreateInstance(&TaskScheduler, None, CLSCTX_ALL)?;
            task_service.Connect(empty_variant(), empty_variant(), empty_variant(), empty_variant())?;

            let task_definition: ITaskDefinition = task_service.NewTask(0)?;
            let triggers: ITriggerCollection = task_definition.get_Triggers()?;
            let reg_info: IRegistrationInfo = task_definition.get_RegistrationInfo()?;
            let actions: IActionCollection = task_definition.get_Actions()?;
            let settings: ITaskSettings = task_definition.get_Settings()?;

            Ok(
                Self {
                    name,
                    task_service,
                    task_definition,
                    reg_info,
                    triggers,
                    actions,
                    settings,
                    folder: None,
                }
            )
        }
    }

    pub fn register(self) -> Result<(), windows::Error> {
        unsafe {
            self.folder.unwrap().RegisterTaskDefinition(
                to_bstr(&self.name),
                self.task_definition,
                TASK_CREATE_OR_UPDATE.0,
                None,
                None,
                TASK_LOGON_INTERACTIVE_TOKEN,
                None
            )?;
            self.settings.put_Enabled(1)?;
        }
        Ok(())
    }

    pub fn set_hidden(&mut self, is_hidden: bool) -> Result<&mut Self, windows::Error> {
        let hidden: i16 = if is_hidden {1} else {0};
        unsafe {
            self.settings.put_Hidden(hidden)?;
        }
        Ok(self)
    }

    pub fn folder(&mut self, path: &str) -> Result<&mut Self, windows::Error> {
        self.folder = Some(
            unsafe {
                self.task_service.GetFolder(to_bstr(path))?
            }
        );
        Ok(self)
    }

    pub fn author(&mut self, author: &str) -> Result<&mut Self, windows::Error> {
        unsafe {
            self.reg_info.put_Author(to_bstr(author))?;
        }
        Ok(self)
    }

    pub fn description(&mut self, description: &str) -> Result<&mut Self, windows::Error> {
        unsafe {
            self.reg_info.put_Description(to_bstr(description))?;
        }
        Ok(self)
    }

    // fn hidden(&mut self, is_hidden: bool) -> Result<&mut Self, windows::Error> {
    //
    // }

    pub fn idle_trigger(&mut self, idle_trigger: TaskIdleTrigger) -> Result<&mut Self, windows::Error> {
        unsafe {
            let trigger = self.triggers.Create(TASK_TRIGGER_IDLE)?;

            let i_idle_trigger = trigger.cast::<IIdleTrigger>()?;
            i_idle_trigger.put_Id(idle_trigger.id)?;
            i_idle_trigger.put_Enabled(1)?;

            let repetition: IRepetitionPattern = i_idle_trigger.get_Repetition()?;
            repetition.put_Interval(idle_trigger.repetition_interval)?;
            repetition.put_StopAtDurationEnd(idle_trigger.repetition_stop_at_duration_end)?;
        }
        Ok(self)
    }

    pub fn logon_trigger(&mut self, logon_trigger: TaskLogonTrigger) -> Result<&mut Self, windows::Error> {
        unsafe {
            let trigger = self.triggers.Create(TASK_TRIGGER_LOGON)?;
            let i_idle_trigger = trigger.cast::<ILogonTrigger>()?;
            i_idle_trigger.put_Id(logon_trigger.id)?;
            i_idle_trigger.put_Enabled(1)?;

            let repetition = i_idle_trigger.get_Repetition()?;
            repetition.put_Interval(logon_trigger.repetition_interval)?;
            repetition.put_StopAtDurationEnd(logon_trigger.repetition_stop_at_duration_end)?;

            i_idle_trigger.put_Delay(logon_trigger.delay)?;
        }
        Ok(self)
    }

    pub fn principal(&mut self, run_level: RunLevel, id: &str, user_id: &str) -> Result<&mut Self, windows::Error> {
        unsafe {
            let principal = self.task_definition.get_Principal()?;
            match run_level {
                RunLevel::HIGHEST => principal.put_RunLevel(TASK_RUNLEVEL_HIGHEST)?
            }
            principal.put_Id(to_bstr(id))?;
            principal.put_UserId(to_bstr(user_id))?;
        }
        Ok(self)
    }

    pub fn settings(&mut self, task_settings: TaskSettings) -> Result<&mut Self, windows::Error> {
        unsafe {
            self.settings.put_RunOnlyIfIdle(task_settings.run_only_if_idle)?;
            self.settings.put_WakeToRun(task_settings.wake_to_run)?;
            self.settings.put_ExecutionTimeLimit(task_settings.execution_time_limit)?;
            self.settings.put_DisallowStartIfOnBatteries(task_settings.disallow_start_if_on_batteries)?;

            let idle_s = self.settings.get_IdleSettings()?;
            idle_s.put_StopOnIdleEnd(task_settings.idle_settings.stop_on_idle_end)?;
            idle_s.put_RestartOnIdle(task_settings.idle_settings.restart_on_idle)?;
            idle_s.put_IdleDuration(task_settings.idle_settings.idle_duration)?;
            idle_s.put_WaitTimeout(task_settings.idle_settings.wait_timeout)?;
        }

        Ok(self)
    }

    pub fn exec_action(&mut self, task_action: TaskAction) -> Result<&mut Self, windows::Error> {
        unsafe {
            let action: IAction = self.actions.Create(TASK_ACTION_EXEC)?;
            let exec_action: IExecAction = action.cast()?;

            exec_action.put_Path(task_action.path)?;
            exec_action.put_Id(task_action.id)?;
            exec_action.put_WorkingDirectory(task_action.working_dir)?;
            exec_action.put_Arguments(task_action.args)?;
        }
        Ok(self)
    }

    pub fn get_task(&self, task_id: &str) -> Result<IRegisteredTask, windows::Error> {
        unsafe {
            self.folder.as_ref().unwrap().GetTask(to_bstr(task_id))
        }
    }

}

