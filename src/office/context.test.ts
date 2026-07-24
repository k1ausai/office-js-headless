import { describe, expect, it } from "vitest";
import { createOfficeGlobal } from "./context";

describe("createOfficeGlobal", () => {
  it("Office.context throws before onReady() has resolved", () => {
    const office = createOfficeGlobal("PC");
    expect(() => office.context).toThrow();
  });

  it("Office.context.platform reports PC/Mac/OfficeOnline correctly after onReady()", async () => {
    for (const platform of ["PC", "Mac", "OfficeOnline"] as const) {
      const office = createOfficeGlobal(platform);
      await office.onReady();
      expect(office.context.platform).toBe(office.PlatformType[platform]);
    }
  });

  it("isSetSupported('WordApiDesktop', '1.1') is true on PC/Mac, false on OfficeOnline", async () => {
    const pc = createOfficeGlobal("PC");
    await pc.onReady();
    expect(pc.context.requirements.isSetSupported("WordApiDesktop", "1.1")).toBe(true);

    const mac = createOfficeGlobal("Mac");
    await mac.onReady();
    expect(mac.context.requirements.isSetSupported("WordApiDesktop", "1.1")).toBe(true);

    const web = createOfficeGlobal("OfficeOnline");
    await web.onReady();
    expect(web.context.requirements.isSetSupported("WordApiDesktop", "1.1")).toBe(false);
  });

  it("onReady() invokes the optional callback with {host, platform} and also resolves with it", async () => {
    const office = createOfficeGlobal("Mac");
    let callbackInfo: unknown;
    const resolved = await office.onReady((info) => {
      callbackInfo = info;
    });
    expect(resolved).toEqual({ host: office.HostType.Word, platform: office.PlatformType.Mac });
    expect(callbackInfo).toEqual(resolved);
  });

  it("Office.context.requirements.isSetSupported throws before onReady() has resolved, same as .platform", () => {
    const office = createOfficeGlobal("PC");
    expect(() => office.context.requirements).toThrow();
  });
});
