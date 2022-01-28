import {
  ISlideOptionsContext,
  SlideOptionsModule, VueInstance,
} from "dynamicscreen-sdk-js";

export default class MicrosoftDriverOptions extends SlideOptionsModule {
  async onReady() {
    return true;
  };

  setup(props: Record<string, any>, vue: VueInstance, context: ISlideOptionsContext) {
    const { h } = vue;

    return () =>
      h("div")
  }
}
