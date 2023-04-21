use regex::{ Regex, Captures };
use std::collections::HashSet;

pub fn fix_formula(line: &str, futures: &HashSet<&str>) -> String {
    let re = Regex::new(r"([A-Z\.]+)\(").unwrap();
    let fixed = str::replace(line, ";", ",").clone();
                                                        

    let out = re.replace(fixed.as_str(), |caps: &Captures| {
        let op = caps.get(1).unwrap().as_str();
        if futures.contains(op) {
            return format!("_xlfn.{}(", op);
        }
        return caps.get(0).unwrap().as_str().to_string();
    });

    out.into_owned()
}

pub fn get_future_functions() -> HashSet<&'static str>{
    let future: HashSet<&str> = vec!["ACOT", "ACOTH", "AGGREGATE", "ARABIC", "ARRAYTOTEXT", "BASE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "CEILING.MATH", "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "COMBINA", "CONCAT", "CONFIDENCE.NORM", "CONFIDENCE.T", "COT", "COTH", "COVARIANCE.P", "COVARIANCE.S", "CSC", "CSCH", "DAYS", "DECIMAL", "ECMA.CEILING", "ERF.PRECISE", "ERFC.PRECISE", "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FIELDVALUE", "FILTERXML", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY", "FORECAST.ETS.STAT", "FORECAST.LINEAR", "FORMULATEXT", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE", "GAUSS", "HYPGEOM.DIST", "IFNA", "IFS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN", "ISFORMULA", "ISO.CEILING", "ISOWEEKNUM", "LET", "LOGNORM.DIST", "LOGNORM.INV", "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL", "MUNIT", "NEGBINOM.DIST", "NETWORKDAYS.INTL", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NUMBERVALUE", "PDURATION", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUTATIONA", "PHI", "POISSON.DIST", "QUARTILE.EXC", "QUARTILE.INC", "QUERYSTRING", "RANDARRAY", "RANK.AVG", "RANK.EQ", "RRI", "SEC", "SECH", "SEQUENCE", "SHEET", "SHEETS", "SKEW.P", "SORTBY", "STDEV.P", "STDEV.S", "SWITCH", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "TEXTJOIN", "UNICHAR", "UNICODE", "UNIQUE", "VAR.P", "VAR.S", "WEBSERVICE", "WEIBULL.DIST", "WORKDAY.INTL", "XLOOKUP", "XMATCH", "XOR", "Z.TEST"].iter().cloned().collect();
    future
}