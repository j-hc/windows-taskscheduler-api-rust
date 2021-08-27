use crate::variant_d::to_bstr;
use bindings::Windows::Win32::Foundation::BSTR;


pub struct TaskAction {
    pub id: BSTR,
    pub path: BSTR,
    pub working_dir: BSTR,
    pub args: BSTR
}
impl TaskAction {
    pub fn new(id: &str, path: &str, working_dir: &str, args: &str) -> Self {
        Self{
            id: to_bstr(id),
            path: to_bstr(path),
            working_dir: to_bstr(working_dir),
            args: to_bstr(args)
        }
    }
}
