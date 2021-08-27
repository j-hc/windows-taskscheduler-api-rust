use bindings::Windows::Win32::{
    Foundation::BSTR,
    System::{
        OleAutomation::{
            VARIANT, VARIANT_0, VARIANT_0_0_abi, VARIANT_0_0_0_abi, VARIANT_0_0_0_0_abi,
        },
    },
};


use windows::Abi;
use std::ffi::c_void;
use std::ptr;


pub fn to_bstr(str: &str) -> BSTR {
    BSTR::from(str)
}

pub fn to_variant(str: &str) -> VARIANT {
    let bstr: BSTR = to_bstr(str);
    let mut variant = empty_variant();
    variant.Anonymous.Anonymous.vt = 0;
    variant.Anonymous.Anonymous.Anonymous.bstrVal = bstr.abi();
    variant
}

// fn to_int_variant(n: i32) -> VARIANT {
//     let mut variant = def_variant();
//     variant.Anonymous.Anonymous.vt = 0;
//     variant.Anonymous.Anonymous.Anonymous.lVal = n;
//     variant
// }


pub fn empty_variant() -> VARIANT {
    VARIANT {
        Anonymous: VARIANT_0 {
            Anonymous: VARIANT_0_0_abi {
                vt: 0,
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: VARIANT_0_0_0_abi {
                    Anonymous: VARIANT_0_0_0_0_abi {
                        pvRecord: ptr::null_mut() as *mut c_void,
                        pRecInfo: ptr::null_mut() as *mut c_void
                    }
                }
            }
        }
    }
}
