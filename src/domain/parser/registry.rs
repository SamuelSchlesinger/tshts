//! Submodule of `parser` — see parser/mod.rs.

#![allow(unused_imports)]
use super::*;

pub struct FunctionRegistry {
    functions: HashMap<String, FunctionImpl>,
    /// Purity classification per registered function. Defaults to
    /// `Pure` for entries registered via `register_function`; volatile
    /// functions register via `register_function_with_purity`.
    purities: HashMap<String, FunctionPurity>,
}

thread_local! {
    /// One-shot built-in registry per thread. Populating ~140 HashMap entries
    /// on every formula eval was the dominant cost in tight recalc loops; we
    /// now reuse a single registry across calls. Custom functions registered
    /// via the public API still get a fresh registry because they need
    /// independent state.
    static BUILTIN_REGISTRY: std::cell::OnceCell<std::rc::Rc<FunctionRegistry>>
        = const { std::cell::OnceCell::new() };
}

impl FunctionRegistry {
    /// Build a registry with all built-in functions. Use `shared_builtin()`
    /// in hot paths to skip the ~140-entry HashMap rebuild.
    pub fn new() -> Self {
        let mut registry = Self {
            functions: HashMap::new(),
            purities: HashMap::new(),
        };
        registry.register_builtin_functions();
        registry
    }

    /// Shared per-thread singleton for the built-in registry. Cheaper to
    /// clone an `Rc` than to rebuild the HashMap on every formula eval.
    pub fn shared_builtin() -> std::rc::Rc<FunctionRegistry> {
        BUILTIN_REGISTRY.with(|cell| {
            cell.get_or_init(|| std::rc::Rc::new(FunctionRegistry::new()))
                .clone()
        })
    }
    
    /// Registers a new function. Default purity is `Pure`; use
    /// `register_function_with_purity` for volatile / side-effecting
    /// functions.
    pub fn register_function(&mut self, name: &str, func: FunctionImpl) {
        let key = name.to_uppercase();
        self.functions.insert(key.clone(), func);
        // Only set Pure if no prior purity exists — re-registration
        // with the volatile API should win regardless of order.
        self.purities.entry(key).or_insert(FunctionPurity::Pure);
    }

    /// Registers a new function and explicitly tags its purity. Used by
    /// volatile / side-effecting builtins (RAND, NOW, GET, etc.).
    pub fn register_function_with_purity(
        &mut self,
        name: &str,
        func: FunctionImpl,
        purity: FunctionPurity,
    ) {
        let key = name.to_uppercase();
        self.functions.insert(key.clone(), func);
        self.purities.insert(key, purity);
    }

    /// Gets a function by name.
    pub fn get_function(&self, name: &str) -> Option<&FunctionImpl> {
        self.functions.get(&name.to_uppercase())
    }

    /// Look up a function's purity. Returns `Pure` for unknown functions
    /// — the formula walker will surface an error elsewhere if the
    /// function name itself is invalid, so this fallback is safe.
    pub fn purity(&self, name: &str) -> FunctionPurity {
        self.purities
            .get(&name.to_uppercase())
            .copied()
            .unwrap_or(FunctionPurity::Pure)
    }
    
    /// Registers all built-in spreadsheet functions.
    fn register_builtin_functions(&mut self) {
        // Builtin functions are grouped by category. See
        // src/domain/parser/registry_fns/ — one file per category.
        super::registry_fns::date::register(self);
        super::registry_fns::dynamic_array::register(self);
        super::registry_fns::finance::register(self);
        super::registry_fns::info::register(self);
        super::registry_fns::logical::register(self);
        super::registry_fns::lookup::register(self);
        super::registry_fns::numeric::register(self);
        super::registry_fns::string::register(self);
        super::registry_fns::viz::register(self);
        super::registry_fns::web::register(self);
    }
}

impl Default for FunctionRegistry {
    fn default() -> Self {
        Self::new()
    }
}
