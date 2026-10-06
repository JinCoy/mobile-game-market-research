'use client';
import {useEffect,useLayoutEffect,useState} from 'react';
/** Restore local research controls after visiting a source sheet in this tab. */
export function useResearchState<T>(key:string, initial:T) {
  const [state,setState]=useState<T>(initial);
  const [loadedKey,setLoadedKey]=useState<string|null>(null);
  useLayoutEffect(()=>{
    try {
      const stored=sessionStorage.getItem(key);
      const value=stored===null?initial:JSON.parse(stored);
      const valid=Array.isArray(initial)
        ? Array.isArray(value)&&value.length===initial.length&&value.every(v=>typeof v===typeof initial[0])
        : typeof value===typeof initial;
      setState(valid?value:initial);
    } catch {setState(initial);}
    setLoadedKey(key);
    // The key identifies a fixed source version and control default.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  },[key]);
  useEffect(()=>{if(loadedKey!==key)return;try {sessionStorage.setItem(key,JSON.stringify(state));} catch {/* Controls still work when storage is unavailable. */}},[key,state,loadedKey]);
  return [state,setState] as const;
}
