//! Submodule of `models` — see models/mod.rs.
//!
//! Workbook-level dependency graph, replacement for the per-`Spreadsheet`
//! same-sheet graphs plus the `Workbook` cross-sheet propagation layer.
//!
//! Goal: one bidirectional graph over `(SheetId, row, col)` nodes. The
//! parallel calc executor (PR 4) consumes this graph to produce
//! topological levels; the sequential executor (PR 3) walks the same
//! levels in order. Both engines share one source of truth, so a
//! mistake can't make cross-sheet propagation and same-sheet recalc
//! disagree.
//!
//! PR 1 lands the data structures and the rebuild-from-workbook helper,
//! plus level-based topo sort. The runtime recalc path keeps using the
//! existing per-sheet engine; PR 3 swaps it out for a level-walker
//! powered by this graph. The two coexist on PR 1 so tests verify the
//! new graph produces the same dependency closure as the legacy paths.
//!
//! Why SheetId, not sheet name: a sheet name is a `String` (variable
//! length, case-insensitive lookups, mutable via rename). Using it as a
//! HashMap key costs a `String::clone()` per graph edge. SheetId is a
//! 4-byte `Copy` newtype; it's allocated monotonically on sheet creation
//! and never reused, so rename is free (the ID survives) and delete
//! just marks the ID dead. The user-facing sheet_names array maps
//! `SheetId → String` for display.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};

/// Stable, never-reused identifier for a sheet within a workbook.
/// Allocated monotonically by `Workbook::add_sheet` and serialized so
/// the IDs survive save/load round-trips.
#[derive(
    Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize, Deserialize,
)]
pub struct SheetId(pub u32);

/// A cell address within a workbook. The graph's node type.
pub type NodeKey = (SheetId, usize, usize);

/// Bidirectional dependency graph at workbook scope. `dependents[P]`
/// is the set of nodes that reference `P`; `dependencies[X]` is the set
/// of nodes that `X` references. The two maps are kept in sync — every
/// edge in one direction has its inverse in the other.
#[derive(Debug, Clone, Default)]
pub struct WorkbookGraph {
    /// `dependents[P]` = { X : X's formula references P }.
    /// When P changes, every node in `dependents[P]` is dirty.
    pub dependents: HashMap<NodeKey, HashSet<NodeKey>>,
    /// Inverse of `dependents`. `dependencies[X]` = { P : X references P }.
    /// Used to clear stale edges when a cell's formula changes.
    pub dependencies: HashMap<NodeKey, HashSet<NodeKey>>,
}

impl WorkbookGraph {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn is_empty(&self) -> bool {
        self.dependents.is_empty() && self.dependencies.is_empty()
    }

    /// Clear every edge.
    pub fn clear(&mut self) {
        self.dependents.clear();
        self.dependencies.clear();
    }

    /// Register that `node` depends on each cell in `prereqs`. ADDS edges
    /// to any pre-existing set — does NOT remove edges the node had
    /// before. Use [`set_prereqs`] (which calls `unlink_node` first) when
    /// re-registering a node whose formula changed; that's the safe API.
    /// `link` is kept for incremental adds (rare; mostly used by tests).
    pub fn link(&mut self, node: NodeKey, prereqs: impl IntoIterator<Item = NodeKey>) {
        let mut added_any = false;
        for p in prereqs {
            // Self-edges are silently dropped — they form trivial 1-cycles
            // that the cycle detector would reject, but our cycle detection
            // happens elsewhere (the formula-entry check). The graph
            // remains a strict DAG for non-iterative recalc.
            if p == node {
                continue;
            }
            if self.dependencies.entry(node).or_default().insert(p) {
                added_any = true;
            }
            self.dependents.entry(p).or_default().insert(node);
        }
        if !added_any {
            // No edges added; clean up empty entries we may have created.
            if let Some(s) = self.dependencies.get(&node)
                && s.is_empty()
            {
                self.dependencies.remove(&node);
            }
        }
    }

    /// Replace `node`'s outgoing edges with `prereqs`. Equivalent to
    /// `unlink_node(node); link(node, prereqs)` but as one safe call.
    /// Prefer this over `link` when re-registering after a formula
    /// change — `link` alone leaks the old edges.
    pub fn set_prereqs(&mut self, node: NodeKey, prereqs: impl IntoIterator<Item = NodeKey>) {
        self.unlink_node(node);
        self.link(node, prereqs);
    }

    /// Remove every outgoing edge from `node` (and the matching inverse
    /// entries in `dependents`). Use before re-linking when a cell's
    /// formula changes, or when the cell is deleted entirely.
    pub fn unlink_node(&mut self, node: NodeKey) {
        if let Some(prereqs) = self.dependencies.remove(&node) {
            for p in prereqs {
                if let Some(set) = self.dependents.get_mut(&p) {
                    set.remove(&node);
                    if set.is_empty() {
                        self.dependents.remove(&p);
                    }
                }
            }
        }
    }

    /// Forget every edge touching `node` in either direction. Use when a
    /// cell is removed (clears outgoing edges via `unlink_node`) AND the
    /// node should no longer be a precedent (clears incoming edges so
    /// stale dependents stop pointing at it). Each observer's entry in
    /// `dependencies` shrinks by exactly the edge to `node`; entries
    /// that go empty are removed. Observers themselves remain in the
    /// graph and continue to be discoverable as dependents of any other
    /// precedents they had.
    pub fn forget_node(&mut self, node: NodeKey) {
        self.unlink_node(node);
        if let Some(observers) = self.dependents.remove(&node) {
            for o in observers {
                if let Some(deps) = self.dependencies.get_mut(&o) {
                    deps.remove(&node);
                    if deps.is_empty() {
                        self.dependencies.remove(&o);
                    }
                }
            }
        }
    }

    /// Transitive closure of dependents from `seeds`. Order is unspecified
    /// (HashSet iteration); callers that need a topological order should
    /// use `topo_levels_from_seeds` instead.
    pub fn transitive_dependents(&self, seeds: &HashSet<NodeKey>) -> HashSet<NodeKey> {
        let mut visited: HashSet<NodeKey> = HashSet::new();
        let mut queue: VecDeque<NodeKey> = seeds.iter().copied().collect();
        while let Some(n) = queue.pop_front() {
            if !visited.insert(n) {
                continue;
            }
            if let Some(downs) = self.dependents.get(&n) {
                for &d in downs {
                    if !visited.contains(&d) {
                        queue.push_back(d);
                    }
                }
            }
        }
        visited
    }

    /// Topological levels of the dirty closure starting from `seeds`.
    /// Level 0 = seeds; level k+1 depends only on nodes in levels ≤k.
    /// Each level is internally independent — every cell in level k can be
    /// evaluated in parallel without coordinating with peers, since by
    /// construction no cell in level k depends on another cell in level k.
    ///
    /// Implementation: Kahn's algorithm restricted to the subgraph induced
    /// by `transitive_dependents(seeds) ∪ seeds`. We compute in-degrees
    /// inside that subgraph (edges leaving it are ignored — those are
    /// precedents outside the closure, which are already up to date).
    /// Cells with in-degree 0 form level 0; we then peel them off and
    /// recompute in-degrees, repeating until empty.
    ///
    /// Returns the levels as a `Vec<Vec<NodeKey>>`. Within a level, the
    /// node order is unspecified (the caller can sort if a deterministic
    /// schedule is required for reproducibility — useful for tests).
    ///
    /// Cycle handling: if the induced subgraph contains a cycle, the
    /// Kahn walk stops before consuming every node. The unprocessed nodes
    /// are returned as a final "cycle" level (sorted) so the executor
    /// can hand them to the iterative-calc fallback. Callers that need
    /// strict acyclic guarantees should consult `is_cyclic_in_closure`
    /// first.
    pub fn topo_levels_from_seeds(&self, seeds: &HashSet<NodeKey>) -> TopoLevels {
        let closure = self.transitive_dependents(seeds);
        if closure.is_empty() {
            return TopoLevels { levels: Vec::new(), cyclic: Vec::new() };
        }
        // In-degrees restricted to the closure.
        let mut in_degree: HashMap<NodeKey, usize> =
            closure.iter().map(|&n| (n, 0)).collect();
        for &n in &closure {
            if let Some(deps) = self.dependencies.get(&n) {
                for &p in deps {
                    if closure.contains(&p) && p != n {
                        *in_degree.entry(n).or_insert(0) += 1;
                    }
                }
            }
        }
        let mut levels: Vec<Vec<NodeKey>> = Vec::new();
        let mut current: Vec<NodeKey> =
            in_degree.iter().filter_map(|(k, &d)| if d == 0 { Some(*k) } else { None }).collect();
        let mut processed = 0usize;
        while !current.is_empty() {
            processed += current.len();
            let next_level: Vec<NodeKey> = {
                let mut next: HashSet<NodeKey> = HashSet::new();
                for &n in &current {
                    in_degree.remove(&n);
                    if let Some(downs) = self.dependents.get(&n) {
                        for &d in downs {
                            if let Some(deg) = in_degree.get_mut(&d) {
                                *deg -= 1;
                                if *deg == 0 {
                                    next.insert(d);
                                }
                            }
                        }
                    }
                }
                next.into_iter().collect()
            };
            levels.push(current);
            current = next_level;
        }
        // Anything left in in_degree is part of a cycle (or downstream of
        // a cycle). Sort for determinism so test output is stable.
        let mut cyclic: Vec<NodeKey> = in_degree.into_keys().collect();
        cyclic.sort();
        // Invariant: every node in the closure ended up either placed in
        // a level or marked cyclic.
        debug_assert_eq!(
            processed + cyclic.len(),
            closure.len(),
            "topo + cyclic must cover the entire closure"
        );
        TopoLevels { levels, cyclic }
    }

    /// True if any cycle exists in the closure rooted at `seeds`. Used by
    /// callers that want to switch to the iterative-calc fallback before
    /// running the parallel executor.
    #[allow(dead_code)]
    pub fn is_cyclic_in_closure(&self, seeds: &HashSet<NodeKey>) -> bool {
        !self.topo_levels_from_seeds(seeds).cyclic.is_empty()
    }
}

/// Result of a topological sort over a dirty closure. `levels[i]` lists
/// nodes whose precedents are all in `levels[0..i]`. `cyclic` lists nodes
/// that participate in (or are downstream of) a cycle and could not be
/// placed in any level — these are handed to iterative-calc.
#[derive(Debug, Clone, Default)]
pub struct TopoLevels {
    pub levels: Vec<Vec<NodeKey>>,
    pub cyclic: Vec<NodeKey>,
}

impl TopoLevels {
    /// Total node count across all levels (excluding cyclic remainder).
    pub fn linear_count(&self) -> usize {
        self.levels.iter().map(|l| l.len()).sum()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn n(s: u32, r: usize, c: usize) -> NodeKey { (SheetId(s), r, c) }

    #[test]
    fn link_and_unlink_round_trip() {
        let mut g = WorkbookGraph::new();
        g.link(n(0, 1, 0), [n(0, 0, 0), n(0, 0, 1)]);
        assert!(g.dependencies[&n(0, 1, 0)].contains(&n(0, 0, 0)));
        assert!(g.dependents[&n(0, 0, 0)].contains(&n(0, 1, 0)));
        g.unlink_node(n(0, 1, 0));
        assert!(!g.dependencies.contains_key(&n(0, 1, 0)));
        assert!(!g.dependents.contains_key(&n(0, 0, 0)));
    }

    #[test]
    fn self_edges_dropped() {
        let mut g = WorkbookGraph::new();
        g.link(n(0, 1, 0), [n(0, 1, 0)]);
        assert!(g.is_empty());
    }

    #[test]
    fn topo_levels_linear_chain() {
        // A -> B -> C : dirty A, expect levels = [[A], [B], [C]]
        let a = n(0, 0, 0);
        let b = n(0, 1, 0);
        let c = n(0, 2, 0);
        let mut g = WorkbookGraph::new();
        g.link(b, [a]); // B depends on A
        g.link(c, [b]); // C depends on B
        let seeds: HashSet<NodeKey> = [a].into_iter().collect();
        let levels = g.topo_levels_from_seeds(&seeds);
        assert_eq!(levels.levels.len(), 3);
        assert_eq!(levels.levels[0], vec![a]);
        assert_eq!(levels.levels[1], vec![b]);
        assert_eq!(levels.levels[2], vec![c]);
        assert!(levels.cyclic.is_empty());
    }

    #[test]
    fn topo_levels_diamond() {
        // A -> {B, C} -> D : level 0 = A, level 1 = {B, C}, level 2 = D
        let a = n(0, 0, 0);
        let b = n(0, 1, 0);
        let c = n(0, 1, 1);
        let d = n(0, 2, 0);
        let mut g = WorkbookGraph::new();
        g.link(b, [a]);
        g.link(c, [a]);
        g.link(d, [b, c]);
        let seeds: HashSet<NodeKey> = [a].into_iter().collect();
        let levels = g.topo_levels_from_seeds(&seeds);
        assert_eq!(levels.levels.len(), 3);
        assert_eq!(levels.levels[0], vec![a]);
        assert_eq!(levels.levels[1].len(), 2);
        assert!(levels.levels[1].contains(&b));
        assert!(levels.levels[1].contains(&c));
        assert_eq!(levels.levels[2], vec![d]);
    }

    #[test]
    fn topo_levels_cycle_detected() {
        // A -> B -> A : level walk can't proceed
        let a = n(0, 0, 0);
        let b = n(0, 1, 0);
        let mut g = WorkbookGraph::new();
        g.link(a, [b]);
        g.link(b, [a]);
        let seeds: HashSet<NodeKey> = [a].into_iter().collect();
        let levels = g.topo_levels_from_seeds(&seeds);
        // Both nodes have in-degree 1 inside the closure → nothing in level 0.
        assert!(levels.levels.is_empty());
        assert_eq!(levels.cyclic.len(), 2);
        assert!(g.is_cyclic_in_closure(&seeds));
    }

    #[test]
    fn topo_levels_cross_sheet() {
        // Sheet0!A -> Sheet1!A -> Sheet0!B
        let a0 = n(0, 0, 0);
        let a1 = n(1, 0, 0);
        let b0 = n(0, 0, 1);
        let mut g = WorkbookGraph::new();
        g.link(a1, [a0]);
        g.link(b0, [a1]);
        let seeds: HashSet<NodeKey> = [a0].into_iter().collect();
        let levels = g.topo_levels_from_seeds(&seeds);
        assert_eq!(levels.levels.len(), 3);
        assert_eq!(levels.levels[0], vec![a0]);
        assert_eq!(levels.levels[1], vec![a1]);
        assert_eq!(levels.levels[2], vec![b0]);
    }

    #[test]
    fn transitive_dependents_explores_all_downstream() {
        // A -> B -> C, A -> D
        let a = n(0, 0, 0);
        let b = n(0, 0, 1);
        let c = n(0, 0, 2);
        let d = n(0, 0, 3);
        let mut g = WorkbookGraph::new();
        g.link(b, [a]);
        g.link(c, [b]);
        g.link(d, [a]);
        let seeds: HashSet<NodeKey> = [a].into_iter().collect();
        let closure = g.transitive_dependents(&seeds);
        assert_eq!(closure.len(), 4);
        for &n in &[a, b, c, d] {
            assert!(closure.contains(&n));
        }
    }

    #[test]
    fn forget_node_clears_both_directions() {
        let a = n(0, 0, 0);
        let b = n(0, 0, 1);
        let c = n(0, 0, 2);
        let mut g = WorkbookGraph::new();
        g.link(b, [a]); // B depends on A
        g.link(c, [a]); // C depends on A
        g.forget_node(a);
        assert!(!g.dependents.contains_key(&a));
        assert!(!g.dependencies.contains_key(&b));
        assert!(!g.dependencies.contains_key(&c));
    }
}
