import { createRequire } from 'node:module';
import { chai } from 'vitest';
import chaiDatetime from 'chai-datetime';
import chaiXml from 'chai-xml';

globalThis.verquire = createRequire(import.meta.url)('../utils/verquire');

chai.use(chaiXml);
chai.use(chaiDatetime);
