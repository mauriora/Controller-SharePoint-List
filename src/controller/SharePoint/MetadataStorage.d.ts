// causes webpack to fail -> import { defaultMetadataStorage } from 'class-transformer/types/storage';
// declared the below to use -> import { defaultMetadataStorage } from 'class-transformer/esm5/storage';
declare module 'class-transformer/esm5/storage' {
    import type { MetadataStorage } from 'class-transformer/types/MetadataStorage';
  
    export const defaultMetadataStorage: MetadataStorage;
  }