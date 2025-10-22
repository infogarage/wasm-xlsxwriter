use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;
use crate::wrapper::{ format::Format, formula::Formula };

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatBlank {
    pub(crate) inner: xlsx::ConditionalFormatBlank,
}

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatFormula {
    pub(crate) inner: xlsx::ConditionalFormatFormula,
}

#[wasm_bindgen]
impl ConditionalFormatBlank {
    #[wasm_bindgen(constructor, skip_jsdoc)]
    pub fn new() -> ConditionalFormatBlank {
        ConditionalFormatBlank {
            inner: xlsx::ConditionalFormatBlank::new(),
        }
    }

    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(&*format.lock());
    }
}

#[wasm_bindgen]
impl ConditionalFormatFormula {
    #[wasm_bindgen(constructor, skip_jsdoc)]
    pub fn new() -> ConditionalFormatFormula {
        ConditionalFormatFormula {
            inner: xlsx::ConditionalFormatFormula::new(),
        }
    }

    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(&*format.lock());
    }

    #[wasm_bindgen(js_name = "setRule", skip_jsdoc)]
    pub fn set_rule(&mut self, rule: &Formula) {
        self.inner = self.inner.clone().set_rule(&*rule.lock());
    }
}
