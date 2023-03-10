import * as React from 'react'
import { Spinner, SpinnerSize } from '@fluentui/react'

export interface SplashProps {
  logo: string
  message: string
  title: string
}

export default class Splash extends React.Component<SplashProps> {
  render() {
    const { logo, message, title } = this.props

    return (
      <section className="ms-welcome__progress ms-u-fadeIn500">
        <img width="90" height="90" src={logo} alt={title} title={title} />
        <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{title}</h1>
        <Spinner size={SpinnerSize.large} label={message} />
      </section>
    )
  }
}
