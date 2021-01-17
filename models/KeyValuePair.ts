export class KeyValuePair<T,V>{
    id: T;
    value: V


    constructor(_key:T,_value:V)
    {
        this.id = _key;
        this.value = _value;
    }
}
