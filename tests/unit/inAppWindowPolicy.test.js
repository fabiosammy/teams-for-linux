'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const {
  buildInAppHosts,
  isInAppPopup,
} = require('../../app/mainAppWindow/inAppWindowPolicy');
const options = require('../../app/config/options');

const OUTLOOK = 'https://outlook.office.com/mail/inbox';

function hosts(...entries) {
  return buildInAppHosts(entries);
}

describe('buildInAppHosts', () => {
  it('keeps bare hostnames, lowercased and trimmed', () => {
    assert.deepStrictEqual(
      [...buildInAppHosts([' Outlook.Office.com ', 'outlook.cloud.microsoft'])],
      ['outlook.office.com', 'outlook.cloud.microsoft']
    );
  });

  it('drops anything that is not a bare hostname', () => {
    const warn = console.warn;
    console.warn = () => {};
    try {
      assert.deepStrictEqual(
        [
          ...buildInAppHosts([
            'https://outlook.office.com',
            'outlook.office.com/mail',
            '*.office.com',
            'outlook.office.com:443',
            '.office.com',
            'office..com',
            '',
            42,
            null,
          ]),
        ],
        []
      );
    } finally {
      console.warn = warn;
    }
  });

  it('returns an empty set for a missing or non-array value', () => {
    assert.strictEqual(buildInAppHosts(undefined).size, 0);
    assert.strictEqual(buildInAppHosts('outlook.office.com').size, 0);
  });

  it('defaults to no in-app hosts', () => {
    assert.deepStrictEqual(options.urlHandling.default.inAppHosts, []);
  });
});

describe('isInAppPopup', () => {
  const outlook = hosts('outlook.office.com');

  it('keeps an about:blank popup opened by an in-app host page (pop-out compose)', () => {
    assert.strictEqual(isInAppPopup(outlook, 'about:blank', OUTLOOK), true);
    assert.strictEqual(isInAppPopup(outlook, 'about:blank#pop', OUTLOOK), true);
  });

  it('keeps a popup to an in-app host (open in new window)', () => {
    assert.strictEqual(
      isInAppPopup(outlook, 'https://outlook.office.com/mail/deeplink/read/abc', OUTLOOK),
      true
    );
  });

  it('leaves about:blank from any other opener to the existing handling', () => {
    assert.strictEqual(
      isInAppPopup(outlook, 'about:blank', 'https://teams.cloud.microsoft/'),
      false
    );
    assert.strictEqual(isInAppPopup(outlook, 'about:blank', 'https://ads.example.com/'), false);
    assert.strictEqual(isInAppPopup(outlook, 'about:blank', ''), false);
  });

  it('never keeps a popup Chromium already blocked', () => {
    assert.strictEqual(isInAppPopup(outlook, 'about:blank#blocked', OUTLOOK), false);
  });

  it('sends other hosts, look-alikes and plain http to the existing handling', () => {
    for (const url of [
      'https://example.com/doc',
      'https://login.microsoftonline.com/common/oauth2',
      'https://evil.outlook.office.com/',
      'https://outlook.office.com.evil.com/',
      'http://outlook.office.com/mail/',
      'javascript:alert(1)',
      'not a url',
    ]) {
      assert.strictEqual(isInAppPopup(outlook, url, OUTLOOK), false, url);
    }
  });

  it('keeps nothing when no in-app hosts are configured (Teams default)', () => {
    const none = hosts();
    assert.strictEqual(isInAppPopup(none, 'about:blank', OUTLOOK), false);
    assert.strictEqual(isInAppPopup(none, OUTLOOK, OUTLOOK), false);
  });
});
