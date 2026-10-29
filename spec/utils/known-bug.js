// Registers a test for a known lib bug: the body states the CORRECT behaviour, so it fails today.
// The test passes while the body throws an Error whose message matches `reason`, and fails when
// the bug gets fixed (no error) or breaks differently (other message). Prefer this to it.fails,
// which would also pass on any unrelated breakage.
function expectKnownFailure(fn, reason) {
  if (!(reason instanceof RegExp)) {
    throw new TypeError('known bug tests need a RegExp describing the expected failure');
  }
  return async () => {
    let error;
    try {
      await fn();
    } catch (e) {
      error = e;
    }
    expect(error, 'expected the known bug to still occur').to.be.an('error');
    expect(error.message).to.match(reason);
  };
}

function knownBug(title, fn, reason) {
  it(title, expectKnownFailure(fn, reason));
}

module.exports = { knownBug, expectKnownFailure };
