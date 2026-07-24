// Matches the real Office.PlatformType/HostType enums' numeric values
// (node_modules/@types/office-js/index.d.ts) exactly — unlike
// Word.InsertLocation, these are plain numeric enums, not string-valued.
export const PlatformType = {
  PC: 0,
  OfficeOnline: 1,
  Mac: 2,
  iOS: 3,
  Android: 4,
  Universal: 5,
} as const;

export const HostType = {
  Word: 0,
} as const;

export type SupportedPlatform = "PC" | "Mac" | "OfficeOnline";

const PLATFORM_TYPE_BY_NAME: Record<SupportedPlatform, number> = {
  PC: PlatformType.PC,
  Mac: PlatformType.Mac,
  OfficeOnline: PlatformType.OfficeOnline,
};

// Scoped to the one signal the driving consumer's code actually branches on
// — WordApiDesktop, true on PC/Mac, false on OfficeOnline (a genuine
// capability gap, not a shim limitation). Anything else queried (e.g. plain
// WordApi, supported everywhere in scope) defaults to true rather than
// modeling every real api-set/version pair. `minVersion` is accepted (real
// Office.js's signature takes it) but not consulted — nothing in scope
// queries a version boundary, only WordApiDesktop's binary PC/Mac-vs-web gap.
export function isSetSupported(
  platform: SupportedPlatform,
  name: string,
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  minVersion?: string
): boolean {
  if (name === "WordApiDesktop") {
    return platform !== "OfficeOnline";
  }
  return true;
}

interface OfficeContext {
  platform: number;
  requirements: { isSetSupported(name: string, minVersion?: string): boolean };
}

export interface OfficeGlobal {
  readonly PlatformType: typeof PlatformType;
  readonly HostType: typeof HostType;
  readonly context: OfficeContext;
  onReady(
    callback?: (info: { host: number; platform: number }) => unknown
  ): Promise<{ host: number; platform: number }>;
}

export function createOfficeGlobal(platform: SupportedPlatform): OfficeGlobal {
  let readyContext: OfficeContext | undefined;

  return {
    PlatformType,
    HostType,
    get context(): OfficeContext {
      if (!readyContext) {
        // Real Office.context genuinely isn't populated until the runtime
        // signals readiness — this throws the same way accessing a property
        // on `undefined` would in the real host, not a shim-specific error.
        throw new Error("Office.context is not available until Office.onReady() has resolved");
      }
      return readyContext;
    },
    onReady(callback) {
      return Promise.resolve().then(() => {
        const platformType = PLATFORM_TYPE_BY_NAME[platform];
        readyContext = {
          platform: platformType,
          requirements: {
            isSetSupported: (name: string, minVersion?: string) =>
              isSetSupported(platform, name, minVersion),
          },
        };
        const info = { host: HostType.Word, platform: platformType };
        callback?.(info);
        return info;
      });
    },
  };
}
