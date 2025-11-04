use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;
use crate::wrapper::{ color::Color, format::Format, formula::Formula, datetime::ExcelDateTime };

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

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatDataBar {
    pub(crate) inner: xlsx::ConditionalFormatDataBar,
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

    #[wasm_bindgen(js_name = "setStopIfTrue", skip_jsdoc)]
    pub fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

#[wasm_bindgen]
impl ConditionalFormatDataBar {
    #[wasm_bindgen(constructor, skip_jsdoc)]
    pub fn new() -> ConditionalFormatDataBar {
        ConditionalFormatDataBar {
            inner: xlsx::ConditionalFormatDataBar::new(),
        }
    }

    #[wasm_bindgen(js_name = "setMinimum", skip_jsdoc)]
    pub fn set_minimum(&mut self, rule_type: ConditionalFormatType, value: &ConditionalFormatValue) {
        self.inner = self.inner.clone().set_minimum(rule_type.into(), value.inner.clone());
    }

    #[wasm_bindgen(js_name = "setMaximum", skip_jsdoc)]
    pub fn set_maximum(&mut self, rule_type: ConditionalFormatType, value: &ConditionalFormatValue) {
        self.inner = self.inner.clone().set_maximum(rule_type.into(), value.inner.clone());
    }

    #[wasm_bindgen(js_name = "setFillColor", skip_jsdoc)]
    pub fn set_fill_color(&mut self, color: Color) {
        self.inner = self.inner.clone().set_fill_color(color.inner);
    }

    #[wasm_bindgen(js_name = "setBorderColor", skip_jsdoc)]
    pub fn set_border_color(&mut self, color: Color) {
        self.inner = self.inner.clone().set_border_color(color.inner);
    }

    #[wasm_bindgen(js_name = "setNegativeFillColor", skip_jsdoc)]
    pub fn set_negative_fill_color(&mut self, color: Color) {
        self.inner = self.inner.clone().set_negative_fill_color(color.inner);
    }

    #[wasm_bindgen(js_name = "setNegativeBorderColor", skip_jsdoc)]
    pub fn set_negative_border_color(&mut self, color: Color) {
        self.inner = self.inner.clone().set_negative_border_color(color.inner);
    }

    #[wasm_bindgen(js_name = "setSolidFill", skip_jsdoc)]
    pub fn set_solid_fill(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_solid_fill(enable);
    }

    #[wasm_bindgen(js_name = "setBorderOff", skip_jsdoc)]
    pub fn set_border_off(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_border_off(enable);
    }

    #[wasm_bindgen(js_name = "setDirection", skip_jsdoc)]
    pub fn set_direction(&mut self, direction: ConditionalFormatDataBarDirection) {
        self.inner = self.inner.clone().set_direction(direction.into());
    }

    #[wasm_bindgen(js_name = "setBarOnly", skip_jsdoc)]
    pub fn set_bar_only(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_bar_only(enable);
    }

    #[wasm_bindgen(js_name = "setAxisPosition", skip_jsdoc)]
    pub fn set_axis_position(&mut self, position: ConditionalFormatDataBarAxisPosition) {
        self.inner = self.inner.clone().set_axis_position(position.into());
    }

    #[wasm_bindgen(js_name = "setAxisColor", skip_jsdoc)]
    pub fn set_axis_color(&mut self, color: Color) {
        self.inner = self.inner.clone().set_axis_color(color.inner);
    }
}

#[derive(Debug, Clone, Copy, Default)]
#[wasm_bindgen]
pub enum ConditionalFormatType {
    #[default]
    Automatic,
    Lowest,
    Number,
    Percent,
    Formula,
    Percentile,
    Highest,
}

impl From<ConditionalFormatType> for xlsx::ConditionalFormatType {
    fn from(value: ConditionalFormatType) -> xlsx::ConditionalFormatType {
        match value {
            ConditionalFormatType::Automatic => xlsx::ConditionalFormatType::Automatic,
            ConditionalFormatType::Lowest => xlsx::ConditionalFormatType::Lowest,
            ConditionalFormatType::Number => xlsx::ConditionalFormatType::Number,
            ConditionalFormatType::Percent => xlsx::ConditionalFormatType::Percent,
            ConditionalFormatType::Formula => xlsx::ConditionalFormatType::Formula,
            ConditionalFormatType::Percentile => xlsx::ConditionalFormatType::Percentile,
            ConditionalFormatType::Highest => xlsx::ConditionalFormatType::Highest,
        }
    }
}

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatValue {
    pub(crate) inner: xlsx::ConditionalFormatValue,
}

#[wasm_bindgen]
impl ConditionalFormatValue {
    #[wasm_bindgen(js_name = "fromString")]
    pub fn from_string(value: &str) -> ConditionalFormatValue {
        ConditionalFormatValue { inner: xlsx::ConditionalFormatValue::from(value) }
    }

    #[wasm_bindgen(js_name = "fromNumber")]
    pub fn from_number(num: f64) -> ConditionalFormatValue {
        ConditionalFormatValue { inner: xlsx::ConditionalFormatValue::from(num) }
    }

    #[wasm_bindgen(js_name = "fromBool")]
    pub fn from_bool(val: bool) -> ConditionalFormatValue {
        ConditionalFormatValue { inner: xlsx::ConditionalFormatValue::from(val) }
    }

    // Formula型から生成
    #[wasm_bindgen(js_name = "fromFormula")]
    pub fn from_formula(formula: &Formula) -> ConditionalFormatValue {
        ConditionalFormatValue { inner: xlsx::ConditionalFormatValue::from(formula.lock().clone()) }
    }

    // ExcelDateTime型から生成
    #[wasm_bindgen(js_name = "fromExcelDateTime")]
    pub fn from_excel_date_time(dt: &ExcelDateTime) -> ConditionalFormatValue {
        ConditionalFormatValue { inner: xlsx::ConditionalFormatValue::from(dt.inner.lock().unwrap().clone()) }
    }
}

#[derive(Clone)]
#[wasm_bindgen]
pub enum ConditionalFormatDataBarDirection {
    Context,
    LeftToRight,
    RightToLeft,
}

impl From<ConditionalFormatDataBarDirection> for xlsx::ConditionalFormatDataBarDirection {
    fn from(direction: ConditionalFormatDataBarDirection) -> xlsx::ConditionalFormatDataBarDirection {
        match direction {
            ConditionalFormatDataBarDirection::Context => xlsx::ConditionalFormatDataBarDirection::Context,
            ConditionalFormatDataBarDirection::LeftToRight => xlsx::ConditionalFormatDataBarDirection::LeftToRight,
            ConditionalFormatDataBarDirection::RightToLeft => xlsx::ConditionalFormatDataBarDirection::RightToLeft,
        }
    }
}

#[derive(Clone)]
#[wasm_bindgen]
pub enum ConditionalFormatDataBarAxisPosition {
    Automatic,
    Midpoint,
    None,
}

impl From<ConditionalFormatDataBarAxisPosition> for xlsx::ConditionalFormatDataBarAxisPosition {
    fn from(direction: ConditionalFormatDataBarAxisPosition) -> xlsx::ConditionalFormatDataBarAxisPosition {
        match direction {
            ConditionalFormatDataBarAxisPosition::Automatic => xlsx::ConditionalFormatDataBarAxisPosition::Automatic,
            ConditionalFormatDataBarAxisPosition::Midpoint => xlsx::ConditionalFormatDataBarAxisPosition::Midpoint,
            ConditionalFormatDataBarAxisPosition::None => xlsx::ConditionalFormatDataBarAxisPosition::None,
        }
    }
}
