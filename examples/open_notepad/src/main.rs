use std::time::Duration;
use windows_taskscheduler::task_action::TaskAction;
use windows_taskscheduler::{
    task::{RunLevel, Task},
    task_trigger,
};

fn main() -> windows_taskscheduler::Result<()> {
    let trigger = task_trigger::TaskIdleTrigger::new(
        "idletrigger",
        Duration::from_secs(3 * 60),
        true,
        Duration::from_secs(10 * 60),
    );
    let action = TaskAction::new("action", r"notepad.exe", r"", "");
    let task = Task::new()?
        .folder(r"\")?
        .idle_trigger(trigger)?
        .exec_action(action)?
        .principal(RunLevel::LUA, r"", r"")?
        .set_hidden(false)?;

    task.register("open notepad when idle")?;

    Ok(())
}
