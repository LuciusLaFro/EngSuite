import { DefaultButton, mergeStyles } from '@fluentui/react'
import React from 'react'

export default function Header({ label, tipText, buttonText, onClick, imgSrc, imgAlt}) {
  return (
    <div>
        <div className={headerContainer}>
            <div className={headerWrapper}>
                <h1 className={headerLabel}>{label}</h1>
                <img className={headerImage} src={imgSrc} alt={imgAlt}></img>
            </div>
            <p className={tipParagraph}>{tipText}</p>
            <DefaultButton onClick={onClick} className={button__styles}>{buttonText}</DefaultButton>
        </div>
    </div>
  )
}

const headerContainer = mergeStyles({
    boxShadow: '0px 5px 10px -5px #6D6D6D',
    padding: '0px 10px',
    marginBottom: '7px',
    // background: '#BFBFBF'
})

const headerWrapper = mergeStyles({
    display: 'grid',
    gridTemplateColumns: '5fr 1fr',
    padding: '0px 5px 0px 5px',
    marginBottom: '0px',
    // background: '#6D6D6D'
})

const tipParagraph = mergeStyles({
    padding: '0px 0px 10px 15px',
    marginTop: '0px',
    marginBottom: '0px',
})

const headerLabel = mergeStyles({
    marginLeft: '2.5%',
    marginTop:'0',
    marginBottom: '5px',
})

const headerImage = mergeStyles({
    width: '35px',
    height: '25px',
    margin: '10px',
})

const button__styles = mergeStyles({
    width: "100%",
    marginBottom: "10px",
    marginTop: "0px",
})