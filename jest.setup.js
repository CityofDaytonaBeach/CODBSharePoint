// Define build-time globals that are normally injected by esbuild
globalThis.__VERSION__ = '1.0.0';
globalThis.__BUILD_TIME__ = new Date().toISOString();
