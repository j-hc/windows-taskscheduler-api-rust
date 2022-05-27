# windows task scheduler api for rust

This was made for personal use so know that it's very limited

## Usage
In your Cargo.toml
```rust
windows-taskscheduler = { git = "https://github.com/j-hc/windows-taskscheduler-api-rust.git" }
```

Also have a look at the [example here](/examples/open_notepad.rs)
#
```rust
use std::time::Duration;
use windows_taskscheduler::{TaskAction, RunLevel, Task, TaskIdleTrigger};


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
    .register("open notepad when idle")?;

```
