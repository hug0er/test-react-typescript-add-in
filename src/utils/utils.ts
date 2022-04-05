export const toPromise = <T>(
  callback: (params: unknown, value: (params: T) => void) => void,
  params: unknown
): Promise<T> => {
  const promise = new Promise((resolve, reject) => {
    try {
      params ? callback(params, resolve) : callback(params, resolve);
    } catch (err) {
      reject(err);
    }
  });
  return promise as Promise<T>;
};

export const evaluateIncludesInList = (array: string[], str: string): boolean => {
  for (const item in array) {
    if (str.includes(item)) return true;
  }

  return false;
};
