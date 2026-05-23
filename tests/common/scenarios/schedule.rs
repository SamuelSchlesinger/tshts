//! Project schedule (Gantt-style cascading). Each task has a duration
//! and a list of predecessor tasks. Earliest start of a task = max(end
//! of its predecessors); end = start + duration. Tests `MAX` over
//! arbitrary cell ranges + a multi-level dependency cascade.
//!
//! Project end = max end across all tasks.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Schedule;

#[derive(Clone)]
pub struct Task {
    pub name: &'static str,
    pub duration: f64,
    pub predecessors: Vec<usize>, // 0-indexed indices into tasks
}

#[derive(Clone)]
pub struct Inputs {
    pub tasks: Vec<Task>,
}

pub struct Output {
    pub starts: Vec<f64>,
    pub ends: Vec<f64>,
    pub project_end: f64,
}

impl Scenario for Schedule {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "schedule" }

    fn default_inputs(&self) -> Inputs {
        // A small software project. Topo-sorted so the population loop
        // can rely on predecessors being declared earlier.
        Inputs {
            tasks: vec![
                Task { name: "spec",    duration: 5.0,  predecessors: vec![] },
                Task { name: "ui_mock", duration: 3.0,  predecessors: vec![0] },
                Task { name: "api",     duration: 8.0,  predecessors: vec![0] },
                Task { name: "db",      duration: 4.0,  predecessors: vec![0] },
                Task { name: "fe_impl", duration: 7.0,  predecessors: vec![1, 2] },
                Task { name: "be_impl", duration: 9.0,  predecessors: vec![2, 3] },
                Task { name: "qa",      duration: 4.0,  predecessors: vec![4, 5] },
                Task { name: "launch",  duration: 1.0,  predecessors: vec![6] },
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let n = i.tasks.len();
        let mut starts = vec![0.0; n];
        let mut ends = vec![0.0; n];
        for (idx, t) in i.tasks.iter().enumerate() {
            let s = t.predecessors
                .iter()
                .map(|&p| ends[p])
                .fold(0.0f64, f64::max);
            starts[idx] = s;
            ends[idx] = s + t.duration;
        }
        let project_end = ends.iter().copied().fold(0.0f64, f64::max);
        Output { starts, ends, project_end }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Layout: row 1 header, rows 2.. tasks
        //   A: name | B: duration | C: start | D: end
        enter_cells(h, &[
            ("A1", "name"), ("B1", "duration"), ("C1", "start"), ("D1", "end"),
        ]);
        for (idx, t) in i.tasks.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), t.name);
            enter_cell(h, &format!("B{}", row), &lit(t.duration));
            let start_formula = if t.predecessors.is_empty() {
                "0".to_string()
            } else if t.predecessors.len() == 1 {
                format!("=D{}", 2 + t.predecessors[0])
            } else {
                let refs: Vec<String> = t
                    .predecessors
                    .iter()
                    .map(|&p| format!("D{}", 2 + p))
                    .collect();
                format!("=MAX({})", refs.join(","))
            };
            enter_cell(h, &format!("C{}", row), &start_formula);
            enter_cell(h, &format!("D{}", row), &format!("=C{}+B{}", row, row));
        }
        // Project end.
        let last = 1 + i.tasks.len();
        enter_cell(h, "F1", "project_end");
        enter_cell(h, "G1", &format!("=MAX(D2:D{})", last));
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (idx, (s, e)) in o.starts.iter().zip(&o.ends).enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("start_{}", idx + 1),
                format!("C{}", row), *s));
            v.push(CellCheck::new(format!("end_{}", idx + 1),
                format!("D{}", row), *e));
        }
        v.push(CellCheck::new("project_end", "G1", o.project_end));
        v
    }
}
