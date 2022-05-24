use std::time::Duration;
use windows_taskscheduler::RegisteredTask;
use windows_taskscheduler::TaskAction;
use windows_taskscheduler::{RunLevel, Task, TaskIdleTrigger};

fn main() -> windows_taskscheduler::Result<()> {
    let registered_task = create_task()?;
    let registered_task_from_name = Task::get_task(r"\", "open notepad when idle")?;

    assert_eq!(registered_task.name()?, registered_task_from_name.name()?);

    delete_task()?;

    Ok(())
}

fn delete_task() -> windows_taskscheduler::Result<()> {
    Task::delete_task(r"\", "open notepad when idle")
}

fn create_task() -> windows_taskscheduler::Result<RegisteredTask> {
    let trigger = TaskIdleTrigger::new(
        "idletrigger",
        Duration::from_secs(3 * 60),
        true,
        Duration::from_secs(10 * 60),
    );
    let action = TaskAction::new("action", "notepad.exe", "", "");
    Task::new(r"\")?
        .idle_trigger(trigger)?
        .exec_action(action)?
        .principal(RunLevel::LUA, "", "")?
        .set_hidden(false)?
        .register("open notepad when idle")
}
