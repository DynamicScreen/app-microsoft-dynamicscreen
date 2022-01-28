import {
  ISlideContext,
  SlideModule,
  VueInstance,
} from "dynamicscreen-sdk-js"

export default class Excel extends SlideModule {
  async onReady() {
    return true;
  };

  setup(props: Record<string, any>, vue: VueInstance, context: ISlideContext) {
    const { h, ref, reactive } = vue;

    return () => [
      h('div', 'hello')
      ]
  }
}
