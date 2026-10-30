import * as React from 'react';
import styles from './StyledText.module.scss';
import { IStyledTextProps } from './IStyledTextProps';


const StyledText: React.FC<IStyledTextProps> = (
  props: IStyledTextProps
): React.ReactElement => {

  const normaliseColour = (
    value: string | undefined,
    fallback: string
  ): string => {

    const trimmedValue: string =
      (value || '').trim();

    return trimmedValue.length > 0
      ? trimmedValue
      : fallback;
  };


  const clampNumber = (
    value: number,
    minimum: number,
    maximum: number
  ): number => {

    if (!Number.isFinite(value)) {
      return minimum;
    }

    return Math.min(
      maximum,
      Math.max(minimum, value)
    );
  };


  const getFontFamily = (): string => {

    if (props.fontFamilyChoice === 'custom') {

      const customFont: string =
        (props.customFontFamily || '').trim();

      if (customFont.length > 0) {
        return customFont;
      }

      return '"Segoe UI", Arial, sans-serif';
    }

    return (
      props.fontFamilyChoice ||
      '"Segoe UI", Arial, sans-serif'
    );
  };


  const getBlockMargins = (): {
    marginLeft: string;
    marginRight: string;
  } => {

    switch (props.textBlockPosition) {

      case 'centre':
        return {
          marginLeft: 'auto',
          marginRight: 'auto'
        };

      case 'right':
        return {
          marginLeft: 'auto',
          marginRight: '0'
        };

      case 'left':
      default:
        return {
          marginLeft: '0',
          marginRight: 'auto'
        };
    }
  };


  const convertHexToRgba = (
    colour: string,
    alpha: number
  ): string | undefined => {

    const hexValue: string =
      colour.replace('#', '').trim();

    let red: number;
    let green: number;
    let blue: number;

    if (hexValue.length === 3) {

      red = parseInt(
        hexValue.charAt(0) +
        hexValue.charAt(0),
        16
      );

      green = parseInt(
        hexValue.charAt(1) +
        hexValue.charAt(1),
        16
      );

      blue = parseInt(
        hexValue.charAt(2) +
        hexValue.charAt(2),
        16
      );

    } else if (
      hexValue.length === 6 ||
      hexValue.length === 8
    ) {

      red = parseInt(
        hexValue.substring(0, 2),
        16
      );

      green = parseInt(
        hexValue.substring(2, 4),
        16
      );

      blue = parseInt(
        hexValue.substring(4, 6),
        16
      );

    } else {

      return undefined;
    }

    if (
      Number.isNaN(red) ||
      Number.isNaN(green) ||
      Number.isNaN(blue)
    ) {
      return undefined;
    }

    return (
      `rgba(` +
      `${red}, ` +
      `${green}, ` +
      `${blue}, ` +
      `${alpha}` +
      `)`
    );
  };


  const convertRgbToRgba = (
    colour: string,
    alpha: number
  ): string | undefined => {

    const match: RegExpMatchArray | null =
      colour.match(
        /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*(?:0|1|0?\.\d+))?\s*\)$/i
      );

    if (!match) {
      return undefined;
    }

    const red: number =
      clampNumber(
        Number(match[1]),
        0,
        255
      );

    const green: number =
      clampNumber(
        Number(match[2]),
        0,
        255
      );

    const blue: number =
      clampNumber(
        Number(match[3]),
        0,
        255
      );

    return (
      `rgba(` +
      `${red}, ` +
      `${green}, ` +
      `${blue}, ` +
      `${alpha}` +
      `)`
    );
  };


  const getBackgroundColour = (): string => {

    if (!props.useBackground) {
      return 'transparent';
    }

    const colour: string =
      normaliseColour(
        props.backgroundColour,
        '#FFFFFF'
      );

    const opacityPercentage: number =
      clampNumber(
        props.backgroundOpacity,
        0,
        100
      );

    const alpha: number =
      opacityPercentage / 100;

    if (alpha === 0) {
      return 'transparent';
    }

    if (colour.startsWith('#')) {

      return (
        convertHexToRgba(
          colour,
          alpha
        ) ||
        'transparent'
      );
    }

    if (
      colour
        .toLowerCase()
        .startsWith('rgb')
    ) {

      return (
        convertRgbToRgba(
          colour,
          alpha
        ) ||
        'transparent'
      );
    }

    return (
      `color-mix(` +
      `in srgb, ` +
      `${colour} ${opacityPercentage}%, ` +
      `transparent` +
      `)`
    );
  };


  const fontColour: string =
    normaliseColour(
      props.fontColour,
      '#323130'
    );

  const gradientStartColour: string =
    normaliseColour(
      props.gradientStartColour,
      '#0078d4'
    );

  const gradientMiddleColour: string =
    normaliseColour(
      props.gradientMiddleColour,
      '#6b4eff'
    );

  const gradientEndColour: string =
    normaliseColour(
      props.gradientEndColour,
      '#6b2fa0'
    );


  const getGradientBackground = (): string => {

    if (props.useGradientMiddleColour) {

      return (
        `linear-gradient(` +
        `${props.gradientAngle}deg, ` +
        `${gradientStartColour} 0%, ` +
        `${gradientMiddleColour} 50%, ` +
        `${gradientEndColour} 100%` +
        `)`
      );
    }

    return (
      `linear-gradient(` +
      `${props.gradientAngle}deg, ` +
      `${gradientStartColour} 0%, ` +
      `${gradientEndColour} 100%` +
      `)`
    );
  };


  const textAlign:
    React.CSSProperties['textAlign'] =
      props.textAlignment === 'centre'
        ? 'center'
        : props.textAlignment;


  const blockMargins = getBlockMargins();


  const textWidth: number =
    clampNumber(
      props.textWidth,
      10,
      100
    );

  const horizontalPadding: number =
    props.useBackground
      ? clampNumber(
          props.backgroundPaddingHorizontal,
          0,
          100
        )
      : 0;

  const verticalPadding: number =
    props.useBackground
      ? clampNumber(
          props.backgroundPaddingVertical,
          0,
          100
        )
      : 0;

  const borderRadius: number =
    props.useBackground
      ? clampNumber(
          props.backgroundBorderRadius,
          0,
          100
        )
      : 0;


  const outerStyle: React.CSSProperties = {

    width:
      props.widthMode === 'fitContent'
        ? 'fit-content'
        : `${textWidth}%`,

    maxWidth:
      '100%',

    boxSizing:
      'border-box',

    marginTop:
      `${props.marginTop}px`,

    marginBottom:
      `${props.marginBottom}px`,

    marginLeft:
      blockMargins.marginLeft,

    marginRight:
      blockMargins.marginRight,

    paddingLeft:
      `${horizontalPadding}px`,

    paddingRight:
      `${horizontalPadding}px`,

    paddingTop:
      `${verticalPadding}px`,

    paddingBottom:
      `${verticalPadding}px`,

    borderRadius:
      `${borderRadius}px`,

    backgroundColor:
      getBackgroundColour()
  };


  const innerTextStyle: React.CSSProperties = {

    display:
      'block',

    fontFamily:
      getFontFamily(),

    fontSize:
      `${props.fontSize}px`,

    fontWeight:
      props.fontWeight,

    letterSpacing:
      `${props.letterSpacing}px`,

    lineHeight:
      props.lineHeight,

    textAlign,

    textTransform:
      props.textTransform,

    whiteSpace:
      'pre-wrap',

    overflowWrap:
      'break-word',

    wordBreak:
      'normal'
  };


  if (props.useGradient) {

    innerTextStyle.backgroundImage =
      getGradientBackground();

    innerTextStyle.backgroundClip =
      'text';

    innerTextStyle.WebkitBackgroundClip =
      'text';

    innerTextStyle.color =
      'transparent';

    innerTextStyle.WebkitTextFillColor =
      'transparent';

  } else {

    innerTextStyle.backgroundImage =
      'none';

    innerTextStyle.backgroundClip =
      'border-box';

    innerTextStyle.WebkitBackgroundClip =
      'border-box';

    innerTextStyle.color =
      fontColour;

    innerTextStyle.WebkitTextFillColor =
      fontColour;
  }


  return (
    <div className={styles.styledTextWebPart}>
      <div className={styles.textContainer}>
        <div
          className={styles.styledText}
          style={outerStyle}
        >
          <span style={innerTextStyle}>
            {props.text ||
              'Enter your text in the web part properties.'}
          </span>
        </div>
      </div>
    </div>
  );
};


export default StyledText;