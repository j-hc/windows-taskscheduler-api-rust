use std::time::Duration;
use windows_taskscheduler::RegisteredTask;
use windows_taskscheduler::TaskAction;
use windows_taskscheduler::TaskLogonTrigger;
use windows_taskscheduler::{RunLevel, Task, TaskIdleTrigger};

fn main() -> windows_taskscheduler::Result<()> {
    const TASK_NAME: &str = "open_notepad_when_idle";

    let _registered_task = create_task(TASK_NAME)?;
    let registered_task = Task::get_task(r"\", TASK_NAME)?;

    assert_eq!(registered_task.name()?, registered_task.name()?);

    registered_task.run_raw()?;

    println!("{:?}", registered_task.last_run_time()?);
    println!("{:?}", registered_task.next_run_time()?);

    Task::delete_task(r"\", TASK_NAME)?;

    Ok(())
}

fn create_task(name: &str) -> windows_taskscheduler::Result<RegisteredTask> {
    let idle_trigger = TaskIdleTrigger::new(
        "idletrigger",
        Duration::from_secs(3 * 60),
        true,
        Duration::from_secs(10 * 60),
    );

    // requires admin rights
    let _logon_trigger = TaskLogonTrigger::new(
        "logontrigger",
        Duration::from_secs(3 * 60),
        true,
        Duration::from_secs(10),
        Duration::from_secs(1),
    );

    let action = TaskAction::new("action", "notepad.exe", "", "");
    Task::new(r"\")?
        .idle_trigger(idle_trigger)?
        // .logon_trigger(logon_trigger)?
        .exec_action(action)?
        .principal(RunLevel::LUA, "", "")?
        .set_hidden(false)?
        .register(name)
}
