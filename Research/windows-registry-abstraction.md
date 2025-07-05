# Abstracting Windows Registry Access via a Library

When looking to **abstract Windows Registry interactions** (akin to how
`System.IO.Abstractions` abstracts file I/O), an ideal approach is to use a
dedicated library that provides an interface-based registry wrapper. This avoids
reinventing the wheel with custom interfaces and mock implementations. One such
solution is the **DotNetWindowsRegistry** library by Dave Kerr, which offers a
**test-friendly abstraction layer for the registry**.

## Existing Registry Abstraction Library (DotNetWindowsRegistry)

**DotNetWindowsRegistry** is an open-source NuGet package that wraps the static
`Microsoft.Win32.Registry` API with interface-driven classes. It introduces
interfaces like `IRegistry` and `IRegistryKey` (analogous to the standard
`Registry`/`RegistryKey` classes) and provides two implementations:

- **`WindowsRegistry`** – a concrete implementation that delegates to the real
  Windows Registry (calls the actual Win32 APIs under the hood). This is used in
  production code to perform real registry reads/writes.
- **`InMemoryRegistry`** – an in-memory fake implementation that simulates the
  registry structure in RAM. This is used for unit testing, allowing you to
  verify registry operations without touching the actual OS registry.

Using this library, you **replace direct calls** to `Microsoft.Win32.Registry`
with calls to the `IRegistry` interface. For example, instead of
`Registry.CurrentUser.OpenSubKey(...)` you would call the injected
`IRegistry.OpenBaseKey(...)`, and instead of `RegistryKey.SetValue` you use
`IRegistryKey.SetValue`, etc.. Under the hood, `WindowsRegistry` just forwards
these calls to the real registry, whereas `InMemoryRegistry` stores the data in
a mock registry hive for testing. The library is designed to be API-compatible
with the built-in Registry classes (it’s _“100% compliant with the existing
Microsoft.Win32.Registry package”_), making it straightforward to integrate.

**Benefits of this approach:**

- _No custom coding needed:_ DotNetWindowsRegistry already defines the
  interfaces and provides a robust in-memory simulation. This saves you from
  writing your own wrappers and ensures the abstraction is well-tested by the
  community.
- _Familiar API:_ You can use methods and patterns similar to the standard
  Registry API, minimizing the learning curve and code changes.
- _Ease of testing:_ In unit tests, swap in the `InMemoryRegistry`
  implementation. You can pre-populate it with keys/values and later inspect its
  state to assert that your code made the correct registry changes, all without
  touching the real registry. Dave Kerr’s example shows populating an
  `InMemoryRegistry` with a known structure, executing code, and then calling an
  `.Print()` or similar method to verify the expected keys/values were written.

To get started, you can add the **DotNetWindowsRegistry** package from NuGet and
begin refactoring your code to use `IRegistry`. For instance, if you had a class
writing to the registry, you would inject an `IRegistry` into it (see next
section) and use that for all registry calls. In production you pass a
`WindowsRegistry` instance, and in tests you pass an `InMemoryRegistry` – the
behavior remains the same, but no real registry is touched in tests.

## Enforcing Dependency Injection for Registry Operations

Adopting this abstraction goes hand-in-hand with using **Dependency Injection
(DI)** throughout your codebase for any registry access. The goal is to
**eliminate all direct, static calls** to `Microsoft.Win32.Registry` in favor of
calls via an injected interface (often called `IRegistryService` or just
`IRegistry`). This means:

- **Define a registry interface** (if using the library, `IRegistry` is
  provided; otherwise you might define your own `IRegistryService` with needed
  methods). All code that needs to read/write the registry will depend on this
  interface, not on the static Registry classes.

- **Inject the registry dependency** into your classes. For example:

  ```csharp
  public class FooManager
  {
      private readonly IRegistry _registry;
      public FooManager(IRegistry registry)
      {
          _registry = registry;
      }
      public void SaveFooSetting(string name, string value)
      {
          using var key = _registry.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
          using var subKey = key.CreateSubKey("MyApp\\Settings");
          subKey.SetValue(name, value);
      }
  }
  ```

  Here `FooManager` doesn’t call `Registry.CurrentUser` directly – it uses the
  injected `_registry`. This makes `FooManager` agnostic about the actual
  registry backend. In your **composition root** or during startup, you would
  configure DI so that `IRegistry` is bound to a `WindowsRegistry` instance (the
  real implementation). In unit tests, you’d instantiate `FooManager` with an
  `InMemoryRegistry` or a mock implementation of `IRegistry` as needed.

- **Apply DI consistently** for _all current and future_ components that need
  registry access. By uniformly passing an `IRegistry`/`IRegistryService` to
  each class (via constructor injection, which is recommended), you ensure that
  no part of the code is secretly using the static registry calls. This uniform
  approach is crucial for testability and flexibility. It allows you to swap
  implementations easily – not only for testing (e.g., use the in-memory
  registry) but also if you ever needed to redirect registry calls elsewhere
  (for example, if running on a non-Windows platform or wanting to store
  configuration in a different medium).

This DI-centric design is essentially an application of the Dependency Injection
principle, promoting loose coupling in your architecture. As the
DotNetWindowsRegistry documentation notes, _“the only way to be able to test
changes to the registry with this library is to make sure that your code uses
the IRegistry and IRegistryKey interfaces – not the concrete
Microsoft.Win32.Registry class. In practice this means you will most likely need
to adopt a Dependency Injection pattern for your code.”_ In other words, **every
registry operation goes through the injected interface**, enabling you to
intercept those calls in tests or swap out the backing implementation
effortlessly.

By using a ready-made abstraction library and enforcing DI, you achieve your
goals of improved unit testing, maintainability, and flexibility:

- **Easier Unit Testing:** With DI, you can inject an in-memory or mock registry
  to test how your code handles registry reads/writes, without affecting the
  real registry or requiring special test setup/teardown of actual registry
  keys. This leads to more reliable and isolated tests.
- **Clean Architecture:** The application logic is decoupled from
  platform-specific details. Your classes don’t know _how_ or _where_ the data
  is stored (registry vs. some future alternative); they just call methods on an
  interface. This adheres to clean architecture principles and makes future
  maintenance easier.
- **No Reinventing the Wheel:** By leveraging an existing, well-tested library,
  you avoid the pitfalls of writing your own abstraction from scratch. The
  DotNetWindowsRegistry library was specifically created to solve this problem
  in a reusable way, so you can trust it to handle the nuances of registry
  access while you focus on your application logic.

In summary, you should introduce an **IRegistry abstraction layer** (ideally via
an existing library like _DotNetWindowsRegistry_) and refactor all registry
calls to use this interface through DI. In practice, that means passing an
`IRegistry` (or `IRegistryService`) into any class that needs to read/write the
registry, binding it to a real implementation at runtime and to a fake or
in-memory implementation in tests. This approach gives you the same kind of
flexibility and testability for the Windows Registry that
`System.IO.Abstractions` provides for file systems – allowing you to unit test
registry interactions confidently and maintain a cleaner separation of concerns
in your codebase.

**Sources:**

- Kerr, Dave. “Unit Testing the Windows Registry.” _dwmkerr.com_, 5 Sept 2020.
- **DotNetWindowsRegistry** – GitHub project README (Dave Kerr’s registry
  abstraction library).
