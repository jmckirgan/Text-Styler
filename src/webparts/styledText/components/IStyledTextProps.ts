export type StyledTextAlignment =
  | 'left'
  | 'centre'
  | 'right';

export type StyledTextTransform =
  | 'none'
  | 'uppercase'
  | 'lowercase'
  | 'capitalize';

export type StyledTextBlockPosition =
  | 'left'
  | 'centre'
  | 'right';

export type StyledTextWidthMode =
  | 'percentage'
  | 'fitContent';

export interface IStyledTextProps {
  text: string;

  fontFamilyChoice: string;
  customFontFamily: string;

  fontSize: number;
  fontWeight: number;
  fontColour: string;

  letterSpacing: number;
  lineHeight: number;

  textAlignment: StyledTextAlignment;
  textTransform: StyledTextTransform;

  widthMode: StyledTextWidthMode;
  textWidth: number;
  textBlockPosition: StyledTextBlockPosition;

  useGradient: boolean;
  gradientStartColour: string;

  useGradientMiddleColour: boolean;
  gradientMiddleColour: string;

  gradientEndColour: string;
  gradientAngle: number;

  useBackground: boolean;
  backgroundColour: string;
  backgroundOpacity: number;
  backgroundPaddingHorizontal: number;
  backgroundPaddingVertical: number;
  backgroundBorderRadius: number;

  marginTop: number;
  marginBottom: number;
}