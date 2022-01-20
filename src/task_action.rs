use windows::Win32::Foundation::BSTR;

pub struct TaskAction {
    pub(crate) id: BSTR,
    pub(crate) path: BSTR,
    pub(crate) working_dir: BSTR,
    pub(crate) args: BSTR,
}
impl TaskAction {
    pub fn new(id: &str, path: &str, working_dir: &str, args: &str) -> Self {
        Self {
            id: id.into(),
            path: path.into(),
            working_dir: working_dir.into(),
            args: args.into(),
        }
    }
}
