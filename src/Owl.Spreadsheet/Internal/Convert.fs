namespace Owl.Spreadsheet

module Convert =
  /// <summary></summary>
  let inline to_bool(value: obj) = 
    match value with 
    | :? bool as b -> b
    | :? string as s -> not <| System.String.IsNullOrEmpty(s)
    | :? System.DateTime as date -> System.Convert.ToBoolean(date.Ticks)
    | :? System.TimeSpan as span -> System.Convert.ToBoolean(span.Ticks)
    | _ -> System.Convert.ToBoolean(value)
  /// <summary></summary>
  let inline to_single(value: obj) = 
    match value with 
    | :? string as s -> if System.String.IsNullOrEmpty(s) then 0f else System.Single.Parse(s)
    | :? System.DateTime as date -> System.Convert.ToSingle(date.Ticks)
    | :? System.TimeSpan as span -> System.Convert.ToSingle(span.Ticks)
    | _ -> System.Convert.ToSingle(value)
  /// <summary></summary>
  let inline to_double(value: obj) = 
    match value with
    | :? string as s -> if System.String.IsNullOrEmpty(s) then 0. else System.Double.Parse(s) 
    | :? System.DateTime as date -> System.Convert.ToDouble(date.Ticks)
    | :? System.TimeSpan as span -> System.Convert.ToDouble(span.Ticks)
    | _ -> System.Convert.ToDouble(value)
  /// <summary></summary>
  let inline to_int16(value: obj) =
    match value with
    | :? string as s -> if System.String.IsNullOrEmpty(s) then 0s else System.Int16.Parse(s)
    | :? System.DateTime as date -> System.Convert.ToInt16(date.Ticks)
    | :? System.TimeSpan as span -> System.Convert.ToInt16(span.Ticks)
    | _ -> System.Convert.ToInt16(value)
  /// <summary></summary>
  let inline to_int(value: obj) =
    match value with
    | :? string as s -> if System.String.IsNullOrEmpty(s) then 0 else System.Int32.Parse(s) 
    | :? System.DateTime as date -> System.Convert.ToInt32(date.Ticks)
    | :? System.TimeSpan as span -> System.Convert.ToInt32(span.Ticks)
    | _ -> System.Convert.ToInt32(value)
  /// <summary></summary>
  let inline to_int64(value: obj) =
    match value with
    | :? string as s -> if System.String.IsNullOrEmpty(s) then 0L else System.Int64.Parse(s)
    | :? System.DateTime as date -> date.Ticks
    | :? System.TimeSpan as span -> span.Ticks
    | _ -> System.Convert.ToInt64(value)
  /// <summary></summary>
  let inline to_decimal(value: obj) =
    match value with
    | :? string as s -> if System.String.IsNullOrEmpty(s) then 0m else System.Decimal.Parse(s)
    | :? System.DateTime as date -> System.Convert.ToDecimal(date.Ticks)
    | :? System.TimeSpan as span -> System.Convert.ToDecimal(span.Ticks)
    | _ -> System.Convert.ToDecimal(value)
  /// <summary></summary>
  let inline to_datetime(value: obj) =
    match value with
    | :? string as s -> if System.String.IsNullOrEmpty(s) then System.DateTime.MinValue else System.DateTime.Parse(s)
    | :? System.DateTime as date -> date
    | :? System.TimeSpan as span -> System.DateTime.MinValue + span
    | _ -> System.Convert.ToDateTime(value)
  /// <summary></summary>
  let to_string(value: obj) = match value with :? string as s -> s | _ -> System.Convert.ToString(value)
