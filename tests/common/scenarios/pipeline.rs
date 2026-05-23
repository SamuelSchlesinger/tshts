//! Sales pipeline / weighted forecast. For each deal we have a value
//! and a probability of close. Weighted forecast = value * probability.
//! Total = SUMPRODUCT.
//!
//! Stages bucketed by probability:
//!   ≥ 0.75 commit
//!   ≥ 0.40 best-case
//!   else   pipeline
//!
//! Tests SUMPRODUCT, SUMIF, COUNTIF.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Pipeline;

#[derive(Clone)]
pub struct Deal {
    pub name: &'static str,
    pub value: f64,
    pub probability: f64,
}

#[derive(Clone)]
pub struct Inputs {
    pub deals: Vec<Deal>,
}

pub struct Output {
    pub weighted_per_deal: Vec<f64>,
    pub total_weighted: f64,
    pub commit_count: f64,
    pub commit_sum: f64,
    pub bestcase_sum: f64,
    pub pipeline_sum: f64,
}

impl Scenario for Pipeline {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "pipeline" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            deals: vec![
                Deal { name: "ACME",      value: 120_000.0, probability: 0.90 },
                Deal { name: "Globex",    value:  80_000.0, probability: 0.75 },
                Deal { name: "Initech",   value: 200_000.0, probability: 0.60 },
                Deal { name: "Umbrella",  value:  50_000.0, probability: 0.40 },
                Deal { name: "Stark",     value: 300_000.0, probability: 0.30 },
                Deal { name: "Wayne",     value:  25_000.0, probability: 0.20 },
                Deal { name: "Cyberdyne", value:  90_000.0, probability: 0.10 },
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let weighted: Vec<f64> = i.deals.iter().map(|d| d.value * d.probability).collect();
        let total: f64 = weighted.iter().sum();
        let commit_count = i.deals.iter().filter(|d| d.probability >= 0.75).count() as f64;
        let commit_sum: f64 = i.deals.iter().filter(|d| d.probability >= 0.75)
            .map(|d| d.value).sum();
        let bestcase_sum: f64 = i.deals.iter()
            .filter(|d| d.probability >= 0.40 && d.probability < 0.75)
            .map(|d| d.value).sum();
        let pipeline_sum: f64 = i.deals.iter().filter(|d| d.probability < 0.40)
            .map(|d| d.value).sum();
        Output { weighted_per_deal: weighted, total_weighted: total,
            commit_count, commit_sum, bestcase_sum, pipeline_sum }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        enter_cells(h, &[
            ("A1", "name"), ("B1", "value"), ("C1", "prob"), ("D1", "weighted"),
        ]);
        for (idx, d) in i.deals.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), d.name);
            enter_cell(h, &format!("B{}", row), &lit(d.value));
            enter_cell(h, &format!("C{}", row), &lit(d.probability));
            enter_cell(h, &format!("D{}", row), &format!("=B{}*C{}", row, row));
        }
        let last = 1 + i.deals.len();
        // Totals + bucket aggregates. Use SUMPRODUCT for the total
        // weighted forecast (independent path from SUM(D)), and
        // SUMIF/COUNTIF for the buckets.
        enter_cells(h, &[
            ("F1", "total_weighted"),
            ("F2", "commit_count"),
            ("F3", "commit_sum"),
            ("F4", "bestcase_sum"),
            ("F5", "pipeline_sum"),
        ]);
        enter_cell(h, "G1", &format!("=SUMPRODUCT(B2:B{},C2:C{})", last, last));
        enter_cell(h, "G2", &format!("=COUNTIF(C2:C{},\">=0.75\")", last));
        enter_cell(h, "G3", &format!("=SUMIF(C2:C{},\">=0.75\",B2:B{})", last, last));
        // Best-case = [0.40, 0.75) — express as SUMIF(>=0.40) - SUMIF(>=0.75).
        enter_cell(h, "G4", &format!(
            "=SUMIF(C2:C{l},\">=0.40\",B2:B{l})-SUMIF(C2:C{l},\">=0.75\",B2:B{l})",
            l = last,
        ));
        enter_cell(h, "G5", &format!("=SUMIF(C2:C{l},\"<0.40\",B2:B{l})", l = last));
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (idx, w) in o.weighted_per_deal.iter().enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("weighted_{}", idx + 1),
                format!("D{}", row), *w).with_tolerance(1e-4));
        }
        v.push(CellCheck::new("total_weighted", "G1", o.total_weighted).with_tolerance(1e-3));
        v.push(CellCheck::new("commit_count",   "G2", o.commit_count));
        v.push(CellCheck::new("commit_sum",     "G3", o.commit_sum).with_tolerance(1e-3));
        v.push(CellCheck::new("bestcase_sum",   "G4", o.bestcase_sum).with_tolerance(1e-3));
        v.push(CellCheck::new("pipeline_sum",   "G5", o.pipeline_sum).with_tolerance(1e-3));
        v
    }
}
