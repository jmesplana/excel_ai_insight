import test from 'node:test';
import assert from 'node:assert/strict';
import { ScreenWakeLock } from '../static/js/screen-wake-lock.js';

const flush = () => new Promise(resolve => setImmediate(resolve));
function setup(request) {
    const document = new EventTarget();
    document.visibilityState = 'visible';
    const statuses = [];
    const locks = [];
    const navigator = {wakeLock: {request: request ?? (async type => {
        assert.equal(type, 'screen');
        const lock = new EventTarget();
        lock.released = false;
        lock.release = async () => { lock.released = true; lock.dispatchEvent(new Event('release')); };
        locks.push(lock);
        return lock;
    })}};
    const controller = new ScreenWakeLock({navigator, document, onStatus: value => statuses.push(value)});
    return {controller, navigator, document, statuses, locks};
}

test('holds one lock during work, releases on stop and acquires again on resume', async () => {
    const {controller, locks, statuses} = setup();
    controller.start(); controller.start(); await flush();
    assert.equal(locks.length, 1);
    assert.match(statuses.at(-1), /Keeping your screen awake/);
    controller.stop(); await flush();
    assert.equal(locks[0].released, true);
    controller.start(); await flush();
    assert.equal(locks.length, 2);
    controller.stop();
});

test('reacquires after returning to a visible tab only while processing', async () => {
    const {controller, locks, document, statuses} = setup();
    controller.start(); await flush();
    document.visibilityState = 'hidden';
    await locks[0].release();
    assert.match(statuses.at(-1), /interrupted/);
    document.dispatchEvent(new Event('visibilitychange')); await flush();
    assert.equal(locks.length, 1);
    document.visibilityState = 'visible';
    document.dispatchEvent(new Event('visibilitychange')); await flush();
    assert.equal(locks.length, 2);
    controller.stop();
    document.dispatchEvent(new Event('visibilitychange')); await flush();
    assert.equal(locks.length, 2);
});

test('a request finishing after stop cannot retain a lock or replace a resumed session', async () => {
    const requests = [];
    const {controller, statuses} = setup(() => new Promise(resolve => requests.push(resolve)));
    controller.start(); controller.stop(); controller.start();
    const oldLock = new EventTarget();
    let released = false;
    oldLock.release = async () => { released = true; };
    requests[0](oldLock); await flush();
    assert.equal(released, true);
    assert.match(statuses.at(-1), /Requesting/);
    const newLock = new EventTarget();
    newLock.release = async () => {};
    requests[1](newLock); await flush();
    assert.equal(controller.lock, newLock);
    assert.match(statuses.at(-1), /Keeping/);
    controller.stop();
});

test('unsupported and denied wake protection provide a fallback without failing work', async () => {
    const unsupported = setup();
    delete unsupported.navigator.wakeLock;
    unsupported.controller.start(); await flush();
    assert.match(unsupported.statuses.at(-1), /unavailable/);
    unsupported.controller.stop();
    const denied = setup(async () => { throw new Error('NotAllowedError'); });
    denied.controller.start(); await flush();
    assert.match(denied.statuses.at(-1), /could not start/);
    denied.controller.stop();
});

test('starting hidden waits until the tab becomes visible', async () => {
    const {controller, document, locks} = setup();
    document.visibilityState = 'hidden';
    controller.start(); await flush();
    assert.equal(locks.length, 0);
    document.visibilityState = 'visible';
    document.dispatchEvent(new Event('visibilitychange')); await flush();
    assert.equal(locks.length, 1);
    controller.stop();
});
