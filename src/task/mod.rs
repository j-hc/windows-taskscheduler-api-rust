pub mod task_action;
pub mod task_settings;
pub mod task_trigger;

use crate::task::task_action::TaskAction;
use crate::task::task_settings::TaskSettings;
use crate::RegisteredTask;
use task_trigger::{TaskIdleTrigger, TaskLogonTrigger};

use windows::core::{Interface, Result};
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CLSCTX_ALL, COINIT_MULTITHREADED, VARIANT,
};
use windows::Win32::System::TaskScheduler::{
    IAction, IActionCollection, IExecAction, IIdleSettings, IIdleTrigger, ILogonTrigger,
    IPrincipal, IRegistrationInfo, IRepetitionPattern, ITaskDefinition, ITaskFolder, ITaskService,
    ITaskSettings, ITriggerCollection, TaskScheduler, TASK_ACTION_EXEC, TASK_CREATE_OR_UPDATE,
    TASK_LOGON_INTERACTIVE_TOKEN, TASK_RUNLEVEL_HIGHEST, TASK_RUNLEVEL_LUA, TASK_TRIGGER_IDLE,
    TASK_TRIGGER_LOGON,
};

pub enum RunLevel {
    HIGHEST,
    LUA,
}

pub struct Task {
    task_definition: ITaskDefinition,
    reg_info: IRegistrationInfo,
    triggers: ITriggerCollection,
    actions: IActionCollection,
    settings: ITaskSettings,
    folder: ITaskFolder,
}
impl Task {
    fn get_task_service() -> Result<ITaskService> {
        // im probably leaking com objects memory by not releasing them but meh
        unsafe {
            CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED)?;

            let task_service: ITaskService = CoCreateInstance(&TaskScheduler, None, CLSCTX_ALL)?;
            task_service.Connect(
                VARIANT::default(),
                VARIANT::default(),
                VARIANT::default(),
                VARIANT::default(),
            )?;
            Ok(task_service)
        }
    }

    pub fn new(path: &str) -> Result<Self> {
        unsafe {
            let task_service = Self::get_task_service()?;

            let task_definition: ITaskDefinition = task_service.NewTask(0)?;
            let triggers: ITriggerCollection = task_definition.Triggers()?;
            let reg_info: IRegistrationInfo = task_definition.RegistrationInfo()?;
            let actions: IActionCollection = task_definition.Actions()?;
            let settings: ITaskSettings = task_definition.Settings()?;
            let folder: ITaskFolder = task_service.GetFolder(BSTR::from(path))?;

            Ok(Self {
                task_definition,
                reg_info,
                triggers,
                actions,
                settings,
                folder,
            })
        }
    }

    pub fn from_xml(self, xml: String) -> Result<Self> {
        unsafe {
            let task_service = Self::get_task_service()?;
            let task_definition: ITaskDefinition = task_service.NewTask(0)?;
            task_definition.SetXmlText(BSTR::from(xml))?;
        }
        Ok(self)
    }

    pub fn register(self, name: &str) -> Result<RegisteredTask> {
        unsafe {
            let registered_task = self.folder.RegisterTaskDefinition(
                BSTR::from(name),
                &self.task_definition,
                TASK_CREATE_OR_UPDATE.0,
                None,
                None,
                TASK_LOGON_INTERACTIVE_TOKEN,
                None,
            )?;
            self.settings.SetEnabled(1)?;
            Ok(RegisteredTask { registered_task })
        }
    }

    pub fn set_hidden(self, is_hidden: bool) -> Result<Self> {
        unsafe { self.settings.SetHidden(is_hidden as i16)? }
        Ok(self)
    }

    pub fn author(self, author: &str) -> Result<Self> {
        unsafe { self.reg_info.SetAuthor(BSTR::from(author))? }
        Ok(self)
    }

    pub fn description(self, description: &str) -> Result<Self> {
        unsafe { self.reg_info.SetDescription(BSTR::from(description))? }
        Ok(self)
    }

    pub fn idle_trigger(self, idle_trigger: TaskIdleTrigger) -> Result<Self> {
        unsafe {
            let trigger = self.triggers.Create(TASK_TRIGGER_IDLE)?;

            let i_idle_trigger: IIdleTrigger = trigger.cast::<IIdleTrigger>()?;
            i_idle_trigger.SetId(idle_trigger.id)?;
            i_idle_trigger.SetEnabled(1)?;
            i_idle_trigger.SetExecutionTimeLimit(idle_trigger.execution_time_limit)?;

            let repetition: IRepetitionPattern = i_idle_trigger.Repetition()?;
            repetition.SetInterval(idle_trigger.repetition_interval)?;
            repetition.SetStopAtDurationEnd(idle_trigger.repetition_stop_at_duration_end)?;
        }
        Ok(self)
    }

    pub fn logon_trigger(self, logon_trigger: TaskLogonTrigger) -> Result<Self> {
        unsafe {
            let trigger = self.triggers.Create(TASK_TRIGGER_LOGON)?;
            let i_logon_trigger = trigger.cast::<ILogonTrigger>()?;
            i_logon_trigger.SetId(logon_trigger.id)?;
            i_logon_trigger.SetEnabled(1)?;
            i_logon_trigger.SetExecutionTimeLimit(logon_trigger.execution_time_limit)?;

            let repetition = i_logon_trigger.Repetition()?;
            repetition.SetInterval(logon_trigger.repetition_interval)?;
            repetition.SetStopAtDurationEnd(logon_trigger.repetition_stop_at_duration_end)?;

            i_logon_trigger.SetDelay(logon_trigger.delay)?;
        }
        Ok(self)
    }

    pub fn principal(self, run_level: RunLevel, id: &str, user_id: &str) -> Result<Self> {
        unsafe {
            let principal: IPrincipal = self.task_definition.Principal()?;
            match run_level {
                RunLevel::HIGHEST => principal.SetRunLevel(TASK_RUNLEVEL_HIGHEST)?,
                RunLevel::LUA => principal.SetRunLevel(TASK_RUNLEVEL_LUA)?,
            }
            principal.SetId(BSTR::from(id))?;
            principal.SetUserId(BSTR::from(user_id))?;
        }
        Ok(self)
    }

    pub fn settings(self, task_settings: TaskSettings) -> Result<Self> {
        unsafe {
            self.settings
                .SetRunOnlyIfIdle(task_settings.run_only_if_idle)?;
            self.settings.SetWakeToRun(task_settings.wake_to_run)?;
            self.settings
                .SetExecutionTimeLimit(task_settings.execution_time_limit)?;
            self.settings
                .SetDisallowStartIfOnBatteries(task_settings.disallow_start_if_on_batteries)?;

            if let Some(idle_settings) = task_settings.idle_settings {
                let idle_s: IIdleSettings = self.settings.IdleSettings()?;
                idle_s.SetStopOnIdleEnd(idle_settings.stop_on_idle_end)?;
                idle_s.SetRestartOnIdle(idle_settings.restart_on_idle)?;
                idle_s.SetIdleDuration(idle_settings.idle_duration)?;
                idle_s.SetWaitTimeout(idle_settings.wait_timeout)?;
            }
        }
        Ok(self)
    }

    pub fn exec_action(self, task_action: TaskAction) -> Result<Self> {
        unsafe {
            let action: IAction = self.actions.Create(TASK_ACTION_EXEC)?;
            let exec_action: IExecAction = action.cast()?;

            exec_action.SetPath(task_action.path)?;
            exec_action.SetId(task_action.id)?;
            exec_action.SetWorkingDirectory(task_action.working_dir)?;
            exec_action.SetArguments(task_action.args)?;
        }
        Ok(self)
    }

    pub fn get_task(path: &str, name: &str) -> Result<RegisteredTask> {
        unsafe {
            let task_service = Self::get_task_service()?;
            let folder = task_service.GetFolder(BSTR::from(path))?;
            let registered_task = folder.GetTask(BSTR::from(name))?;
            Ok(RegisteredTask { registered_task })
        }
    }

    pub fn delete_task(path: &str, name: &str) -> Result<()> {
        unsafe {
            let task_service = Self::get_task_service()?;
            let folder = task_service.GetFolder(BSTR::from(path))?;
            folder.DeleteTask(BSTR::from(name), 0)?;
        }
        Ok(())
    }
}
