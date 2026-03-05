declare module "qrcode" {
  interface QRCodeOptions {
    margin?: number;
    width?: number;
    color?: {
      dark?: string;
      light?: string;
    };
  }

  function toCanvas(
    canvasElement: HTMLCanvasElement,
    text: string,
    options?: QRCodeOptions,
  ): Promise<void>;

  const QRCode: {
    toCanvas: typeof toCanvas;
  };

  export default QRCode;
}
