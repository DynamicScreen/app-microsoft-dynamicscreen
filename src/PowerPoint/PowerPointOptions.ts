import {
  ISlideOptionsContext,
  SlideOptionsModule,
  VueInstance,
} from "dynamicscreen-sdk-js"

export default class PowerPointOptions extends SlideOptionsModule {
  async onReady() {
    return true;
  };

  setup(props: Record<string, any>, vue: VueInstance, context: ISlideOptionsContext) {
    const { h, ref, reactive } = vue;

    return () => [
      h('div', 'hello options')
    ]
  }
}
