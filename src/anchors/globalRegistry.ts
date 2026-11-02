/**
 * Shared singletons for all bundles on the page.
 *
 * The extension (application customizer) and the web part can be loaded in separate bundles or in
 * different versions. Each bundle would then have its own copy
 * of the modules – and therefore its own anchorServiceManager. So that everyone sees the same state
 * (as in the SharePoint original via the _SPEventManager), the singletons live on window.
 *
 * Increment REGISTRY_KEY on incompatible changes to the shared classes.
 */
const REGISTRY_KEY: string = '__spfxAnchorService_v1';

export function getGlobal<T>(name: string, factory: () => T): T {
  const w: any = typeof window !== 'undefined' ? window : {};
  const registry: { [name: string]: any } = w[REGISTRY_KEY] || (w[REGISTRY_KEY] = {});
  if (!registry[name]) {
    registry[name] = factory();
  }
  return registry[name] as T;
}
