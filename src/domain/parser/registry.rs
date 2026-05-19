//! Submodule of `parser` — see parser/mod.rs.

#![allow(unused_imports)]
use super::*;

pub struct FunctionRegistry {
    functions: HashMap<String, FunctionImpl>,
}

impl FunctionRegistry {
    /// Creates a new function registry with built-in functions.
    pub fn new() -> Self {
        let mut registry = Self {
            functions: HashMap::new(),
        };
        
        // Register built-in functions
        registry.register_builtin_functions();
        registry
    }
    
    /// Registers a new function in the registry.
    pub fn register_function(&mut self, name: &str, func: FunctionImpl) {
        self.functions.insert(name.to_uppercase(), func);
    }
    
    /// Gets a function by name.
    pub fn get_function(&self, name: &str) -> Option<&FunctionImpl> {
        self.functions.get(&name.to_uppercase())
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
