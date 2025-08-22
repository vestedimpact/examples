import { Text as PDFText, View } from '@react-pdf/renderer';
import { Style } from '@react-pdf/types';
import { ReactNode } from 'react';

import { AssetImpact } from '../api/types';
import { SDGUtils } from '../utils/sdg';

type ContainerProps = {
  align?: 'flex-start' | 'center' | 'flex-end' | 'space-between';
  children?: ReactNode;
  direction?: 'column' | 'row';
  fixed?: boolean;
  grow?: number;
  id?: string;
  items?: 'flex-start' | 'center' | 'flex-end';
  justify?: 'flex-start' | 'center' | 'flex-end' | 'space-between';
  pageBreak?: boolean;
  style?: Style;
  wrap?: boolean;
};

type TextProps = {
  children: ReactNode;
  color?: string;
  fontSize?: string;
  fontWeight?: 400 | 600;
  lineHeight?: number;
  style?: Style;
  textAlign?: 'left' | 'center' | 'right';
};

const Container = ({
  align = 'center',
  children,
  direction = 'column',
  fixed,
  grow,
  id,
  items,
  justify = 'center',
  pageBreak,
  style = {},
  wrap,
}: ContainerProps) => (
  <View
    break={pageBreak}
    fixed={fixed}
    id={id}
    style={{
      alignContent: align,
      alignItems: items,
      display: 'flex',
      flexDirection: direction,
      flexGrow: grow,
      justifyContent: justify,
      width: '100%',
      ...style,
    }}
    wrap={wrap}
  >
    {children}
  </View>
);

const Text = ({
  children,
  color = '#3F3F3F',
  fontSize = '8pt',
  fontWeight = 400,
  lineHeight = 1.2,
  style = {},
  textAlign = 'left',
}: TextProps) => (
  <PDFText
    style={{
      color,
      fontFamily: 'Poppins',
      fontSize,
      fontWeight,
      lineHeight,
      margin: 0,
      padding: 0,
      textAlign,
      ...style,
    }}
  >
    {children}
  </PDFText>
);

const SeparatorBar = ({ height }: { height: string }) => (
  <Container
    style={{
      borderRight: '1pt dashed black',
      height,
      marginHorizontal: '1pt',
      width: '1pt',
    }}
  />
);

const NegativeImpactBar = ({
  disabled,
  showValue,
  value,
  variant = 'primary',
}: {
  disabled: boolean;
  showValue: boolean;
  value: number;
  variant?: 'primary' | 'secondary';
}) => (
  <Container
    style={{
      backgroundColor: disabled ? '#F2F2F2' : '#DCDCDC',
      borderBottomLeftRadius: '50%',
      borderTopLeftRadius: '50%',
      height: '9pt',
      marginBottom: '1pt',
      marginTop: '3pt',
      width: '100%',
    }}
  >
    {!disabled && showValue && (
      <Container
        style={{
          backgroundColor: variant === 'primary' ? '#3D0F68' : '#7BE9E5',
          borderBottomLeftRadius: '50%',
          borderTopLeftRadius: '50%',
          height: '9pt',
          marginLeft: `${100 - getBarWidth(Math.abs(value))}%`,
          width: `${getBarWidth(Math.abs(value))}%`,
        }}
      >
        <Text
          color={variant === 'primary' ? '#FFFFFF' : undefined}
          fontSize="6pt"
          fontWeight={600}
          style={{ marginLeft: '4pt' }}
        >
          {Math.abs(value).toFixed(0)}
        </Text>
      </Container>
    )}
  </Container>
);

const PositiveImpactBar = ({
  disabled,
  showValue,
  value,
  variant = 'primary',
}: {
  disabled: boolean;
  showValue: boolean;
  value: number;
  variant?: 'primary' | 'secondary';
}) => (
  <Container
    style={{
      backgroundColor: disabled ? '#F2F2F2' : '#DCDCDC',
      borderBottomRightRadius: '50%',
      borderTopRightRadius: '50%',
      height: '9pt',
      marginBottom: '1pt',
      marginTop: '3pt',
      width: '100%',
    }}
  >
    {!disabled && showValue && (
      <Container
        style={{
          backgroundColor: variant === 'primary' ? '#3D0F68' : '#7BE9E5',
          borderBottomRightRadius: '50%',
          borderTopRightRadius: '50%',
          height: '9pt',
          width: `${getBarWidth(Math.abs(value))}%`,
        }}
      >
        <Text
          color={variant === 'primary' ? '#FFFFFF' : undefined}
          fontSize="6pt"
          fontWeight={600}
          style={{ marginRight: '4pt' }}
          textAlign="right"
        >
          {Math.abs(value).toFixed(0)}
        </Text>
      </Container>
    )}
  </Container>
);

export const SDGGoalBarGraph = ({ sdgImpacts }: AssetImpact) => {
  const formattedSDGs = SDGUtils.allGoals.map((goal) => {
    const goalImpact = sdgImpacts.find(({ sdgGoal }) => sdgGoal === goal);
    return {
      label: SDGUtils.getGoalLabelSummary(goal),
      negativeImpact: goalImpact ? goalImpact.negativeImpact : null,
      positiveImpact: goalImpact ? goalImpact.positiveImpact : null,
    };
  });

  return (
    <Container wrap={false}>
      <Text fontSize="6pt" fontWeight={600}>
        Business's negative and positive impacts on SDGs
      </Text>
      <Container direction="row">
        <Container style={{ marginRight: '4pt', marginTop: '6pt' }}>
          <Container direction="row">
            <Container>
              {formattedSDGs.map((impact) => (
                <NegativeImpactBar
                  key={`sdg-goal-graph-negative-${impact.label}`}
                  disabled={impact.negativeImpact === null && impact.positiveImpact === null}
                  showValue={impact.negativeImpact !== null && impact.negativeImpact < 0}
                  value={impact.negativeImpact || 0}
                />
              ))}
            </Container>
            <SeparatorBar height={`${13 * formattedSDGs.length}pt`} />
            <Container>
              {formattedSDGs.map((impact) => (
                <PositiveImpactBar
                  key={`sdg-goal-graph-positve-${impact.label}`}
                  disabled={impact.negativeImpact === null && impact.positiveImpact === null}
                  showValue={impact.positiveImpact !== null && impact.positiveImpact > 0}
                  value={impact.positiveImpact || 0}
                />
              ))}
            </Container>
          </Container>
          <ScaleHeader />
        </Container>
        <Container style={{ marginLeft: '4pt' }}>
          {formattedSDGs.map((impact) => (
            <Container key={`sdg-goal-graph-label-${impact.label}`} style={{ height: '13pt' }}>
              <Text fontSize="6pt" lineHeight={1} style={{ textOverflow: 'ellipsis', overflow: 'hidden' }}>
                {`${impact.label}${impact.negativeImpact === null && impact.positiveImpact === null ? ' (no material impact)' : ''
                  }`}
              </Text>
            </Container>
          ))}
        </Container>
      </Container>
    </Container>
  );
};

export const SDGTargetBarGraph = ({ sdgImpacts }: AssetImpact) => {
  const targets = sdgImpacts.flatMap(({ targetImpacts }) => targetImpacts);

  return (
    <Container wrap={false}>
      <Text fontSize="6pt" fontWeight={600}>
        SDG target impact
      </Text>
      <Container direction="row">
        <Container style={{ marginTop: '6pt', paddingRight: '4pt', width: '50%' }}>
          <Container direction="row">
            <Container>
              {targets.map((impact) => (
                <NegativeImpactBar
                  key={`${title}-graph-negative-${impact.label}`}
                  disabled={impact.negativeImpact === null && impact.positiveImpact === null}
                  showValue={impact.negativeImpact !== null && impact.negativeImpact < 0}
                  value={impact.negativeImpact || 0}
                />
              ))}
            </Container>
            <SeparatorBar height={`${13 * items.length}pt`} />
            <Container>
              {targets.map((impact) => (
                <PositiveImpactBar
                  key={`${title}-graph-positve-${impact.label}`}
                  disabled={impact.negativeImpact === null && impact.positiveImpact === null}
                  showValue={impact.positiveImpact !== null && impact.positiveImpact > 0}
                  value={impact.positiveImpact || 0}
                />
              ))}
            </Container>
          </Container>
          <ScaleHeader />
        </Container>
        <Container style={{ paddingLeft: '4pt', width: '50%' }}>
          {targets.map((item) => (
            <Container key={`${title}-graph-label-${item.label}`} style={{ height: '13pt' }}>
              <Text fontSize="6pt" lineHeight={1} style={{ textOverflow: 'ellipsis', overflow: 'hidden' }}>
                {SDGUtils.getTargetLabel(item.sdgTarget)}
              </Text>
            </Container>
          ))}
        </Container>
      </Container>
    </Container>
  );
};
