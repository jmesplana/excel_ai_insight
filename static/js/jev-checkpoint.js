import {saveCheckpoint} from './run-checkpoint.js';
import {snapshotRun} from './recovery-file.js';

export function checkpoint(action, value, changedIndices = []) {
    return saveCheckpoint('jev', action, value && snapshotRun('jev', value), changedIndices);
}
