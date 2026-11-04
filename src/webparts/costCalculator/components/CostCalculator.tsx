import * as React from 'react';
import { CostSheet } from '../../../cost/CostSheet';
import { CostStore } from '../../../cost/types';

export default function CostCalculator(props: {
  store: CostStore;
  userDisplayName: string;
  theme: 'light' | 'dark';
}): React.ReactElement {
  return <CostSheet store={props.store} userName={props.userDisplayName} theme={props.theme} />;
}
