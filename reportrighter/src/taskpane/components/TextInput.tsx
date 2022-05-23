import { mergeStyles } from '@fluentui/react'
import React from 'react'

export default function TextInput({ className, label, placeholder, onChange}) {
  return (
    <div className={className}>
        <label>{label}</label>
        <input type='text'className={`ms-Grid-col ${cellInput}`} placeholder={placeholder} onChange={onChange}></input>
    </div>
  )
}

const cellInput = mergeStyles({
    left: '3%',
    width: '65%',
})