import { useCallback, useEffect, useRef } from "react";


export function useTimeout(callback, delay) {
    const callbackRef = useRef(callback);
    const timeoutRef: any = useRef();

    useEffect(() => {
        callbackRef.current = callback
    }, [callback]);

    const set = useCallback(() => {
        timeoutRef.current = setTimeout(() => callbackRef.current(), delay)
    }, [delay])

    const clear = useCallback(() => {
        timeoutRef.current && clearTimeout(timeoutRef.current)
    }, [])

    useEffect(() => {
        set()
        return clear
    }, [delay, set, clear])

    const reset = useCallback(() => {
        clear()
        set()
    }, [clear, set])

    return { reset, clear }
}

export function useDebounce(callback, delay, dependencies) {
    const { reset, clear } = useTimeout(callback, delay);
    useEffect(reset, [dependencies, reset])
    useEffect(clear, [])
}

export function useUpdateEffect(callback, dependencies) {
    const firstRenderRef = useRef(true)

    useEffect(() => {
        if(firstRenderRef.current) {
            firstRenderRef.current = false;
            return
        }
        return callback()
    }, dependencies)
}