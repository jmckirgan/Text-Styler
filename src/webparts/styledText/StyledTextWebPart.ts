import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Version } from '@microsoft/sp-core-library';

import {
  type IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import StyledText from './components/StyledText';

import {
  IStyledTextProps,
  StyledTextAlignment,
  StyledTextBlockPosition,
  StyledTextTransform,
  StyledTextWidthMode
} from './components/IStyledTextProps';


export interface IStyledTextWebPartProps {
  text: string;

  /*
   * Legacy property retained so existing web-part
   * instances can migrate to the font dropdown.
   */
  fontFamily?: string;

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


export default class StyledTextWebPart extends
  BaseClientSideWebPart<IStyledTextWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IStyledTextProps> =
      React.createElement(
        StyledText,
        {
          text:
            this.properties.text,

          fontFamilyChoice:
            this.properties.fontFamilyChoice,

          customFontFamily:
            this.properties.customFontFamily,

          fontSize:
            this.properties.fontSize,

          fontWeight:
            this.properties.fontWeight,

          fontColour:
            this.properties.fontColour,

          letterSpacing:
            this.properties.letterSpacing,

          lineHeight:
            this.properties.lineHeight,

          textAlignment:
            this.properties.textAlignment,

          textTransform:
            this.properties.textTransform,

          widthMode:
            this.properties.widthMode,

          textWidth:
            this.properties.textWidth,

          textBlockPosition:
            this.properties.textBlockPosition,

          useGradient:
            this.properties.useGradient,

          gradientStartColour:
            this.properties.gradientStartColour,

          useGradientMiddleColour:
            this.properties.useGradientMiddleColour,

          gradientMiddleColour:
            this.properties.gradientMiddleColour,

          gradientEndColour:
            this.properties.gradientEndColour,

          gradientAngle:
            this.properties.gradientAngle,

          useBackground:
            this.properties.useBackground,

          backgroundColour:
            this.properties.backgroundColour,

          backgroundOpacity:
            this.properties.backgroundOpacity,

          backgroundPaddingHorizontal:
            this.properties.backgroundPaddingHorizontal,

          backgroundPaddingVertical:
            this.properties.backgroundPaddingVertical,

          backgroundBorderRadius:
            this.properties.backgroundBorderRadius,

          marginTop:
            this.properties.marginTop,

          marginBottom:
            this.properties.marginBottom
        }
      );

    ReactDom.render(
      element,
      this.domElement
    );
  }


  protected async onInit(): Promise<void> {

    this._setDefaultProperties();

    await super.onInit();
  }


  protected onDispose(): void {

    ReactDom.unmountComponentAtNode(
      this.domElement
    );
  }


  protected get dataVersion(): Version {

    return Version.parse('1.0');
  }


  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {

    super.onPropertyPaneFieldChanged(
      propertyPath,
      oldValue,
      newValue
    );

    const refreshPropertyPane: boolean =
      (
        propertyPath === 'useGradient' ||
        propertyPath === 'useGradientMiddleColour' ||
        propertyPath === 'fontFamilyChoice' ||
        propertyPath === 'widthMode' ||
        propertyPath === 'useBackground'
      ) &&
      oldValue !== newValue;

    if (refreshPropertyPane) {
      this.context.propertyPane.refresh();
    }

    this.render();
  }


  private _setDefaultProperties(): void {

    this.properties.text ??=
      'Styled text';

    this._migrateLegacyFontFamily();

    this.properties.fontSize ??=
      48;

    this.properties.fontWeight ??=
      700;

    this.properties.fontColour ??=
      '#323130';

    this.properties.letterSpacing ??=
      0;

    this.properties.lineHeight ??=
      1.2;

    this.properties.textAlignment ??=
      'left';

    this.properties.textTransform ??=
      'none';

    this.properties.widthMode ??=
      'percentage';

    this.properties.textWidth ??=
      100;

    this.properties.textBlockPosition ??=
      'left';

    this.properties.useGradient ??=
      false;

    this.properties.gradientStartColour ??=
      '#0078d4';

    this.properties.useGradientMiddleColour ??=
      false;

    this.properties.gradientMiddleColour ??=
      '#6b4eff';

    this.properties.gradientEndColour ??=
      '#6b2fa0';

    this.properties.gradientAngle ??=
      90;

    this.properties.useBackground ??=
      false;

    this.properties.backgroundColour ??=
      '#FFFFFF';

    this.properties.backgroundOpacity ??=
      100;

    this.properties.backgroundPaddingHorizontal ??=
      8;

    this.properties.backgroundPaddingVertical ??=
      4;

    this.properties.backgroundBorderRadius ??=
      8;

    this.properties.marginTop ??=
      0;

    this.properties.marginBottom ??=
      16;
  }


  private _migrateLegacyFontFamily(): void {

    if (this.properties.fontFamilyChoice) {

      this.properties.customFontFamily ??=
        '"Segoe UI", Arial, sans-serif';

      return;
    }

    const legacyFontFamily: string =
      (
        this.properties.fontFamily ||
        ''
      ).trim();

    if (legacyFontFamily.length === 0) {

      this.properties.fontFamilyChoice =
        '"Segoe UI", Arial, sans-serif';

      this.properties.customFontFamily =
        '"Segoe UI", Arial, sans-serif';

      return;
    }

    const matchingStandardFont: string | undefined =
      this._getStandardFontValues().find(
        (fontValue: string) =>
          fontValue.toLowerCase() ===
          legacyFontFamily.toLowerCase()
      );

    if (matchingStandardFont) {

      this.properties.fontFamilyChoice =
        matchingStandardFont;

      this.properties.customFontFamily =
        legacyFontFamily;

      return;
    }

    this.properties.fontFamilyChoice =
      'custom';

    this.properties.customFontFamily =
      legacyFontFamily;
  }


  private _getStandardFontValues(): string[] {

    return [
      '"Aptos", "Segoe UI", Arial, sans-serif',
      '"Aptos Display", "Segoe UI", Arial, sans-serif',
      '"Arial", sans-serif',
      '"Arial Black", Arial, sans-serif',
      '"Courier New", Courier, monospace',
      '"Georgia", serif',
      '"Segoe UI", Arial, sans-serif',
      '"Tahoma", Arial, sans-serif',
      '"Times New Roman", Times, serif',
      '"Trebuchet MS", Arial, sans-serif',
      '"Verdana", Arial, sans-serif'
    ];
  }


  private _getColourErrorMessage(
    value: string
  ): string {

    const trimmedValue: string =
      (value || '').trim();

    if (trimmedValue.length === 0) {

      return (
        'Enter a CSS colour such as #323130.'
      );
    }

    const hexPattern: RegExp =
      /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

    const rgbPattern: RegExp =
      /^rgba?\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}(?:\s*,\s*(?:0|1|0?\.\d+))?\s*\)$/i;

    const hslPattern: RegExp =
      /^hsla?\(.+\)$/i;

    const namedColourPattern: RegExp =
      /^[a-zA-Z]+$/;

    if (
      hexPattern.test(trimmedValue) ||
      rgbPattern.test(trimmedValue) ||
      hslPattern.test(trimmedValue) ||
      namedColourPattern.test(trimmedValue)
    ) {
      return '';
    }

    return (
      'Use a valid CSS colour, for example ' +
      '#0078d4, rgb(0, 120, 212), or navy.'
    );
  }


  private _getCustomFontErrorMessage(
    value: string
  ): string {

    const trimmedValue: string =
      (value || '').trim();

    if (
      this.properties.fontFamilyChoice === 'custom' &&
      trimmedValue.length === 0
    ) {
      return 'Enter at least one CSS font family.';
    }

    return '';
  }


  private _getFontFields():
    IPropertyPaneField<unknown>[] {

    const fontFields:
      IPropertyPaneField<unknown>[] = [

        PropertyPaneDropdown(
          'fontFamilyChoice',
          {
            label:
              'Font family',

            options: [
              {
                key:
                  '"Aptos", "Segoe UI", Arial, sans-serif',

                text:
                  'Aptos'
              },
              {
                key:
                  '"Aptos Display", "Segoe UI", Arial, sans-serif',

                text:
                  'Aptos Display'
              },
              {
                key:
                  '"Arial", sans-serif',

                text:
                  'Arial'
              },
              {
                key:
                  '"Arial Black", Arial, sans-serif',

                text:
                  'Arial Black'
              },
              {
                key:
                  '"Courier New", Courier, monospace',

                text:
                  'Courier New'
              },
              {
                key:
                  '"Georgia", serif',

                text:
                  'Georgia'
              },
              {
                key:
                  '"Segoe UI", Arial, sans-serif',

                text:
                  'Segoe UI'
              },
              {
                key:
                  '"Tahoma", Arial, sans-serif',

                text:
                  'Tahoma'
              },
              {
                key:
                  '"Times New Roman", Times, serif',

                text:
                  'Times New Roman'
              },
              {
                key:
                  '"Trebuchet MS", Arial, sans-serif',

                text:
                  'Trebuchet MS'
              },
              {
                key:
                  '"Verdana", Arial, sans-serif',

                text:
                  'Verdana'
              },
              {
                key:
                  'custom',

                text:
                  'Custom...'
              }
            ]
          }
        )
      ];

    if (
      this.properties.fontFamilyChoice ===
      'custom'
    ) {

      fontFields.push(
        PropertyPaneTextField(
          'customFontFamily',
          {
            label:
              'Custom font family',

            description:
              'Enter a CSS font stack, for example ' +
              '"Montserrat", "Segoe UI", Arial, sans-serif.',

            deferredValidationTime:
              300,

            onGetErrorMessage:
              this._getCustomFontErrorMessage.bind(this)
          }
        )
      );
    }

    return fontFields;
  }


  private _getColourFields():
    IPropertyPaneField<unknown>[] {

    if (!this.properties.useGradient) {

      return [
        PropertyPaneTextField(
          'fontColour',
          {
            label:
              'Text colour',

            description:
              'Enter a CSS colour such as ' +
              '#323130, navy or rgb(50, 49, 48).',

            deferredValidationTime:
              300,

            onGetErrorMessage:
              this._getColourErrorMessage.bind(this)
          }
        )
      ];
    }

    const gradientFields:
      IPropertyPaneField<unknown>[] = [

        PropertyPaneTextField(
          'gradientStartColour',
          {
            label:
              'Gradient start colour',

            deferredValidationTime:
              300,

            onGetErrorMessage:
              this._getColourErrorMessage.bind(this)
          }
        ),

        PropertyPaneToggle(
          'useGradientMiddleColour',
          {
            label:
              'Use middle colour',

            onText:
              'Three-colour gradient',

            offText:
              'Two-colour gradient',

            checked:
              this.properties.useGradientMiddleColour
          }
        )
      ];

    if (
      this.properties.useGradientMiddleColour
    ) {

      gradientFields.push(
        PropertyPaneTextField(
          'gradientMiddleColour',
          {
            label:
              'Gradient middle colour',

            description:
              'This colour is positioned at the ' +
              '50% midpoint of the gradient.',

            deferredValidationTime:
              300,

            onGetErrorMessage:
              this._getColourErrorMessage.bind(this)
          }
        )
      );
    }

    gradientFields.push(
      PropertyPaneTextField(
        'gradientEndColour',
        {
          label:
            'Gradient end colour',

          deferredValidationTime:
            300,

          onGetErrorMessage:
            this._getColourErrorMessage.bind(this)
        }
      ),

      PropertyPaneSlider(
        'gradientAngle',
        {
          label:
            'Gradient angle',

          min:
            0,

          max:
            360,

          step:
            1,

          showValue:
            true
        }
      )
    );

    return gradientFields;
  }


  private _getWidthFields():
    IPropertyPaneField<unknown>[] {

    const widthFields:
      IPropertyPaneField<unknown>[] = [

        PropertyPaneChoiceGroup(
          'widthMode',
          {
            label:
              'Width mode',

            options: [
              {
                key:
                  'percentage',

                text:
                  'Percentage'
              },
              {
                key:
                  'fitContent',

                text:
                  'Fit to text'
              }
            ]
          }
        )
      ];

    if (
      this.properties.widthMode ===
      'percentage'
    ) {

      widthFields.push(
        PropertyPaneSlider(
          'textWidth',
          {
            label:
              'Text width as a percentage',

            min:
              10,

            max:
              100,

            step:
              1,

            showValue:
              true
          }
        )
      );
    }

    return widthFields;
  }


  private _getBackgroundFields():
    IPropertyPaneField<unknown>[] {

    const backgroundFields:
      IPropertyPaneField<unknown>[] = [

        PropertyPaneToggle(
          'useBackground',
          {
            label:
              'Use background',

            onText:
              'Background enabled',

            offText:
              'Transparent',

            checked:
              this.properties.useBackground
          }
        )
      ];

    if (!this.properties.useBackground) {
      return backgroundFields;
    }

    backgroundFields.push(
      PropertyPaneTextField(
        'backgroundColour',
        {
          label:
            'Background colour',

          description:
            'Enter a CSS colour such as #FFFFFF or navy.',

          deferredValidationTime:
            300,

          onGetErrorMessage:
            this._getColourErrorMessage.bind(this)
        }
      ),

      PropertyPaneSlider(
        'backgroundOpacity',
        {
          label:
            'Background opacity',

          min:
            0,

          max:
            100,

          step:
            1,

          showValue:
            true
        }
      ),

      PropertyPaneSlider(
        'backgroundPaddingHorizontal',
        {
          label:
            'Horizontal padding in pixels',

          min:
            0,

          max:
            100,

          step:
            1,

          showValue:
            true
        }
      ),

      PropertyPaneSlider(
        'backgroundPaddingVertical',
        {
          label:
            'Vertical padding in pixels',

          min:
            0,

          max:
            100,

          step:
            1,

          showValue:
            true
        }
      ),

      PropertyPaneSlider(
        'backgroundBorderRadius',
        {
          label:
            'Corner radius in pixels',

          min:
            0,

          max:
            100,

          step:
            1,

          showValue:
            true
        }
      )
    );

    return backgroundFields;
  }


  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    return {
      pages: [
        {
          header: {
            description:
              'Configure the text content and appearance.'
          },

          displayGroupsAsAccordion:
            true,

          groups: [
            {
              groupName:
                'Content',

              isCollapsed:
                false,

              groupFields: [
                PropertyPaneTextField(
                  'text',
                  {
                    label:
                      'Text',

                    description:
                      'Enter the text to display.',

                    multiline:
                      true,

                    rows:
                      6,

                    resizable:
                      true
                  }
                )
              ]
            },

            {
              groupName:
                'Font',

              isCollapsed:
                false,

              groupFields: [
                ...this._getFontFields(),

                PropertyPaneSlider(
                  'fontSize',
                  {
                    label:
                      'Font size in pixels',

                    min:
                      8,

                    max:
                      160,

                    step:
                      1,

                    showValue:
                      true
                  }
                ),

                PropertyPaneDropdown(
                  'fontWeight',
                  {
                    label:
                      'Font weight',

                    options: [
                      {
                        key: 100,
                        text: '100 - Thin'
                      },
                      {
                        key: 200,
                        text: '200 - Extra light'
                      },
                      {
                        key: 300,
                        text: '300 - Light'
                      },
                      {
                        key: 400,
                        text: '400 - Regular'
                      },
                      {
                        key: 500,
                        text: '500 - Medium'
                      },
                      {
                        key: 600,
                        text: '600 - Semi bold'
                      },
                      {
                        key: 700,
                        text: '700 - Bold'
                      },
                      {
                        key: 800,
                        text: '800 - Extra bold'
                      },
                      {
                        key: 900,
                        text: '900 - Black'
                      }
                    ]
                  }
                ),

                PropertyPaneSlider(
                  'letterSpacing',
                  {
                    label:
                      'Letter spacing in pixels',

                    min:
                      -5,

                    max:
                      20,

                    step:
                      0.5,

                    showValue:
                      true
                  }
                ),

                PropertyPaneSlider(
                  'lineHeight',
                  {
                    label:
                      'Line height',

                    min:
                      0.8,

                    max:
                      3,

                    step:
                      0.1,

                    showValue:
                      true
                  }
                ),

                PropertyPaneDropdown(
                  'textTransform',
                  {
                    label:
                      'Text transformation',

                    options: [
                      {
                        key:
                          'none',

                        text:
                          'None'
                      },
                      {
                        key:
                          'uppercase',

                        text:
                          'Uppercase'
                      },
                      {
                        key:
                          'lowercase',

                        text:
                          'Lowercase'
                      },
                      {
                        key:
                          'capitalize',

                        text:
                          'Capitalise each word'
                      }
                    ]
                  }
                )
              ]
            },

            {
              groupName:
                'Colour',

              isCollapsed:
                false,

              groupFields: [
                PropertyPaneToggle(
                  'useGradient',
                  {
                    label:
                      'Use gradient',

                    onText:
                      'Gradient enabled',

                    offText:
                      'Plain colour',

                    checked:
                      this.properties.useGradient
                  }
                ),

                ...this._getColourFields()
              ]
            },

            {
              groupName:
                'Background',

              isCollapsed:
                false,

              groupFields: [
                ...this._getBackgroundFields()
              ]
            },

            {
              groupName:
                'Layout',

              isCollapsed:
                false,

              groupFields: [
                ...this._getWidthFields(),

                PropertyPaneChoiceGroup(
                  'textBlockPosition',
                  {
                    label:
                      'Text block position',

                    options: [
                      {
                        key:
                          'left',

                        text:
                          'Left',

                        iconProps: {
                          officeFabricIconFontName:
                            'AlignLeft'
                        }
                      },
                      {
                        key:
                          'centre',

                        text:
                          'Centre',

                        iconProps: {
                          officeFabricIconFontName:
                            'AlignCenter'
                        }
                      },
                      {
                        key:
                          'right',

                        text:
                          'Right',

                        iconProps: {
                          officeFabricIconFontName:
                            'AlignRight'
                        }
                      }
                    ]
                  }
                ),

                PropertyPaneChoiceGroup(
                  'textAlignment',
                  {
                    label:
                      'Text alignment within the block',

                    options: [
                      {
                        key:
                          'left',

                        text:
                          'Left',

                        iconProps: {
                          officeFabricIconFontName:
                            'AlignLeft'
                        }
                      },
                      {
                        key:
                          'centre',

                        text:
                          'Centre',

                        iconProps: {
                          officeFabricIconFontName:
                            'AlignCenter'
                        }
                      },
                      {
                        key:
                          'right',

                        text:
                          'Right',

                        iconProps: {
                          officeFabricIconFontName:
                            'AlignRight'
                        }
                      }
                    ]
                  }
                ),

                PropertyPaneSlider(
                  'marginTop',
                  {
                    label:
                      'Top margin in pixels',

                    min:
                      0,

                    max:
                      100,

                    step:
                      1,

                    showValue:
                      true
                  }
                ),

                PropertyPaneSlider(
                  'marginBottom',
                  {
                    label:
                      'Bottom margin in pixels',

                    min:
                      0,

                    max:
                      100,

                    step:
                      1,

                    showValue:
                      true
                  }
                )
              ]
            }
          ]
        }
      ]
    };
  }
}