use windows_taskscheduler::{task::{Task, RunLevel}, task_trigger};
use windows_taskscheduler::WinResult;
use windows_taskscheduler::task_action::TaskAction;
use windows_taskscheduler::task_settings::{IdleSettings, TaskSettings};


fn main() -> WinResult<()> {
    let trigger = task_trigger::TaskIdleTrigger::new(
        "idletrigger",
        3,
        true
    );
    let action = TaskAction::new(
        "action",
        r"notepad.exe",
        r"",
        ""
    );

    let mut task = Task::new("opennotepadwhenidlelol".to_owned())?;
    task.folder("\\")?
        .idle_trigger(trigger)?
        .exec_action(action)?
        .principal(RunLevel::HIGHEST, r"NT AUTHORITY\SYSTEM",
                   r"NT AUTHORITY\SYSTEM")?
        .set_hidden(true)?;

    task.register()?;

    Ok(())

}
